import pandas as pd
from glob import glob
from pathlib import PureWindowsPath
import re
from collections import OrderedDict
import argparse
from os.path import isfile
import warnings

# Prepare to accept command line arguments
parser = argparse.ArgumentParser(
    description="""Process microscope excel files to calculate area per"""\
        """ cell and return an output excel file. Output excel file should"""\
        """ be in a different directory than the input files."""
)
parser.add_argument(
    '-i', 
    '--input_path',
    help='File path where excel files are located, can be folder or'\
        ' individual file'
)
parser.add_argument(
    '-o',
    '--output_file',
    help='File path to output excel file'
)
parser.add_argument(
    '-g',
    '--glob',
    action='store_true',
    help='Use glob to find files specified by input_path'
)
parser.add_argument(
    '-r',
    '--recursive',
    action='store_true',
    help='Perform glob search recursively'
)

def spot_area(files,output_file):
    """Takes a list of files or path string and iterates over
    all sheets in all files to calculate the area per cell
    
    Args:
        files (list|str): list of file paths or glob compatible file path string
        output_file (str): output excel file path
    
    Raises:
        TypeError: When the files argument is supplied with the wrong data type
        RunTimeError: When calculations are not performed for any input
        FileExistsError: When output file already exists
        
    >>> experiment_data = spot_area('PATH_TO_FILES','path/my_output_file.xlsx')
    """
    if type(files) == str:
        if files.endswith('xlsx'):
            files = [files]
        else:
            files = glob(files+'/**/*.xlsx',recursive=True)
    elif type(files) == list:
        pass
    else:
        raise TypeError("Input must be a path string or a list of file paths.")
    data = []
    # re pattern for source field to get mouse id and treatments
    source_expr = re.compile(
        r'(?P<SubjectID>[^\s]+)\s+\-\s+(?P<Treat1>[^\s]+)\s+\-\s+(?P<Treat2>[^\s]+).*$'
    )
    for file in files:
        xl = pd.ExcelFile(file)
        path=PureWindowsPath(file)
        # assign the file name to be the experiment name
        experiment = path.stem
        sheets = pd.Series(xl.sheet_names)
        for sheet in sheets:
            df = xl.parse(sheet)
            # get rid of non-ascii characters in column name
            df.rename(
                columns=lambda x: re.sub(r'\s*\[.*\]',r'',x),
                inplace=True,
            )
            # check if the sheet has the correct columns
            required_cols = pd.Series(
                ['Source','BinaryID','NumberObjects','BinaryArea']
            )
            if required_cols.isin(df.columns).all():
                pass
            else:
                warnings.warn(
                    f'Could not parse {sheet} in {file}. Skipping sheet.',
                    UserWarning,
                )
                continue
            # omit the blank lines
            df.dropna(how='all',inplace=True)
            # attempt to parse the source field
            source = df['Source'].iloc[0]
            # handle different human source pattern by forcing it to match mouse
            filtered_source = re.sub(r'Set \d+ \-','',source)
            try:
                parsed_source = OrderedDict(
                    source_expr.match(filtered_source).groupdict()
                )
            except:
                warnings.warn(
                    f'Unable to extract information from source pattern: {source}',
                    UserWarning
                )
                parsed_source = OrderedDict(
                    [('SubjectID','NA'),('Treat1','NA'),('Treat2','NA')]
                )

            # perform computations
            total_area = df.loc[
                df['BinaryID'].str.contains('Threshold'),
                'BinaryArea'
            ].sum()
            num_objects = df.loc[
                df['BinaryID'].str.contains('SpotDetect'),
                'NumberObjects'
            ].sum()
            area_per_cell = total_area/num_objects
            # store results of processing a sheet in an OrderedDict
            expt_result = OrderedDict([
                ('Experiment',experiment),
                ('SheetName',sheet),
                *parsed_source.items(),
                ('Source',source),
                ('TotalArea',total_area),
                ('TotalObjects',num_objects),
                ('AreaPerCell',area_per_cell),
            ])
            data.append(expt_result)
    # check to make sure an output was successfully produced
    if data == []:
        raise RuntimeError(
            'Failed to parse any of the sheets in the input files, or failed'\
                ' to detect any input files.'
        )
    # convert list of OrderedDicts to pandas dataframe
    data = pd.DataFrame(data)
    # export to excel
    if isfile(output_file):
        raise FileExistsError(
            "The output file already exists. Choose a new filename."
        )
    else:
        data.to_excel(output_file,index=False)
    return(data)

if __name__ == "__main__":
    args = parser.parse_args()
    # handle flags relating to glob usage for file searching
    # This allows the user to search for files using glob directly
    if args.glob:
        files = glob(args.input_path,recursive=args.recursive)
        data = spot_area(files,args.output_file)
    else:
        data = spot_area(args.input_path,args.output_file)
    print(
        f'Successfully created output {args.output_file}'
    )