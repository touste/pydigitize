#!/usr/bin/env python3
"""pydigitize.

Usage:
    scan.py [options] [OUTPUT]

Examples:
    scan.py out/
    scan.py out/document.pdf
    scan.py out/ -n document

Args:
    OUTPUT         This can either be a filename or a directory name.

Options:
    -h --help      Show this help.
    --version      Show version.

    -n NAME        Text that will be incorporated into the filename.

    -d DEVICE      Set the device.
    -r RESOLUTION  Set the resolution [default: 300].
    -c PAGES       Page count to scan [default: all pages from ADF]
    
    --no-shrink    Do not shrink resulting pdf.
    --nowait       When scanning multiple pages (with the -c parameter), don't
                   wait for manual confirmation but scan as fast as the scanner
                   can process the pages.

    --verbose      Verbose output
    --debug        Debug output

"""
import datetime
import glob
import logging
import os.path
import re
import sys
import tempfile

import docopt
from sh import cd, mv

try:
    from sh import scanimage
except ImportError:
    print('Error: scanimage command not found. Please install sane.')
    sys.exit(1)

try:
    from sh import tiffcp, tiff2pdf
except ImportError:
    print('Error: tiffcp / tiff2pdf commands not found. Please install libtiff.')
    sys.exit(1)

try:
    from sh import gs
except ImportError:
    print('Error: gs commands not found. Please install ghostscript.')
    sys.exit(1)


logger = logging.getLogger('pydigitize')


VALID_RESOLUTIONS = (100, 200, 300, 400, 600)


def prefix():
    duration = (datetime.datetime.now() - START_TIME).total_seconds()
    return '\033[92m\033[1m+\033[0m [{0:>5.2f}s] '.format(duration)


class Scan:

    def __init__(self, *,
        resolution,
        device,
        output,
        name: str = None,
        datestring: str = None,
        count: int = None,
        nowait: bool = False
    ):
        """
        Initialize scan class.

        Added attributes:

        - resolution
        - device
        - output_path
        - count

        """
        # Validate and store resolution
        def _invalid_res():
            print('Invalid resolution. Please use one of {!r}.'.format(VALID_RESOLUTIONS))
            sys.exit(1)
        try:
            if int(resolution) not in VALID_RESOLUTIONS:
                _invalid_res()
        except ValueError:
            _invalid_res()
        else:
            self.resolution = resolution

        # Store device
        self.device = device

        # set timestamp
        Timestamp = START_TIME.strftime('%Y%m%d') + 'Z'
        
        # Validate and store output path
        if os.path.isdir(output):
            if name is None:
                filename = '{}.pdf'.format(timestamp)
            else:
                filename = name
            output_path = os.path.join(output, filename)
        elif os.path.dirname(output) == '' or os.path.isdir(os.path.dirname(output)):
            output_path = output
        else:
            print('Output directory "{}" must already exist.'.format(output))
            sys.exit(1)
        self.output_path = os.path.abspath(output_path)
        logger.debug('Output path: %s', self.output_path)

        # Store page count
        self.count = count
        self.nowait = nowait

    def prepare_directories(self):
        """
        Prepare the temporary output directories.

        Added attributes:

        - workdir

        """
        print(prefix() + 'Creating temporary directory...')
        self.workdir = tempfile.mkdtemp(prefix='pydigitize-')

    def scan_pages(self):
        """
        Scan pages using ``scanimage``.
        """
        def _scan_page(number: int = None):
            if number is None:
                print(prefix() + 'Scanning all pages...')
            else:
                print(prefix() + 'Scanning page %d/%d...' % (number + 1, self.count))
            scanimage_args = {
                'x': 210, 'y': 297,
                'batch': 'out%d.tif',
                'batch-start': '1000',  # Avoid issues with sorting (e.g. out10 < out2)
                'format': 'tiff',
                'resolution': self.resolution,
                '_ok_code': [0, 7],
            }
            if self.device is not None:
                scanimage_args['device_name'] = self.device
            if number is not None:
                scanimage_args['batch-start'] = number
                scanimage_args['batch-count'] = 1
            logger.debug('Scanimage args: %r' % scanimage_args)

            scanimage(**scanimage_args)

        if self.count:
            for i in range(self.count):
                _scan_page(i)
                if not self.nowait and i < (self.count - 1):
                    try:
                        msg = 'Press <ENTER> to scan page %d (or <CTRL+C> to abort)'
                        input(prefix() + msg % (i + 2))
                    except KeyboardInterrupt:
                        print()
                        print(prefix() + 'Aborting.')
                        sys.exit(1)
        else:
            _scan_page(None)

    def combine_tiffs(self):
        """
        Combine tiffs into single multi-page tiff.
        """
        print(prefix() + 'Combining image files...')
        files = sorted(glob.glob('out*.tif'))
        logger.debug('Joining %r', files)
        tiffcp(files, 'output.tif', c='lzw')

    def convert_tiff_to_pdf(self):
        """
        Convert tiff to pdf.

        TODO: use convert instead?

        """
        print(prefix() + 'Converting to PDF...')
        tiff2pdf('output.tif', p='A4', o='output.pdf')
        
    def shrink_pdf(self):
        """
        Shrink pdf.
        """
        print(prefix() + 'Shrinking PDF...')
        gs_args = ['q', 'dNOPAUSE', 'dBATCH', 'dSAFER', 'sDEVICE=pdfwrite', 
                   'dCompatibilityLevel=1.3', 'dPDFSETTINGS=/screen', 
                   'dEmbedAllFonts=true', 'dSubsetFonts=true', 
                   'dColorImageDownsampleType=/Bicubic', 'dColorImageResolution=185',
                   'dGrayImageDownsampleType=/Bicubic', 'dGrayImageResolution=185',
                   'dMonoImageDownsampleType=/Bicubic', 'dMonoImageResolution=185']
        gs_args.extend(['sOutputFile=output.pdf', 'clean.pdf'])
        gs(*args)
        
    def process(self, *, no_shrink=False):
        # Prepare directories
        self.prepare_directories()
        cd(self.workdir)

        # Scan pages
        self.scan_pages()

        # Combine tiffs into single multi-page tiff
        self.combine_tiffs()

        # Convert tiff to pdf
        self.convert_tiff_to_pdf()

        # Shrink
        if no_shrink is False:
            self.shrink_pdf()
            filename = 'clean.pdf'
        else:
            filename = 'output.pdf'

        # Move file
        print(prefix() + 'Moving resulting file...')
        cd('..')
        mv('{}/{}'.format(self.workdir, filename), self.output_path)

        print('\nDone: %s' % self.output_path)


if __name__ == '__main__':
    args = docopt.docopt(__doc__, version='pydigitize 0.1')
    if args['--debug']:
        logging.basicConfig(level=logging.DEBUG)
    elif args['--verbose']:
        logging.basicConfig(level=logging.INFO)
    else:
        logging.basicConfig(level=logging.WARNING)

    logger.debug('Command line args: %r' % args)

    default_output = tempfile.mkdtemp(prefix='pydigitize-', suffix='-out')

    # Default args
    kwargs = {
        'output': default_output,
    }
    no_shrink = False
    
    # Arguments
    kwargs['resolution'] = args['-r']
    kwargs['device'] = args['-d']
    if args['OUTPUT']:
        kwargs['output'] = args['OUTPUT']
    if args['--no-shrink'] is True:
        no_shrink = True
    if args['-n']:
        kwargs['name'] = args['-n']
    if args['-c']:
        if args['-c'] == 'all pages from ADF':
            kwargs['count'] = None
        else:
            try:
                kwargs['count'] = int(args['-c'])
            except ValueError:
                print('Invalid argument to "-c": %r -> must be numeric!' % args['-c'])
                sys.exit(1)
    kwargs['nowait'] = args['--nowait']

    print('                           ____')
    print('  ________________________/ O  \___/')
    print(' <_/_\_/_\_/_\_/_\_/_\_/_______/   \\\n')

    scan = Scan(**kwargs)
    scan.process(no_shrink=no_shrink)
