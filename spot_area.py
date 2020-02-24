import pandas as pd
from glob import glob
from pathlib import PureWindowsPath
import re
from collections import OrderedDict
import argparse
from os.path import isfile

parser = argparse.ArgumentParser(
    description="""Process microscope excel files to calculate area per cell"""\
        """ and return an output excel file. Output excel file should be in a"""\
        """ different directory than the input files."""
)
parser.add_argument(
    '-i', 
    '--input_path',
    help='file path where excel files are located'
)
parser.add_argument(
    '-o',
    '--output_file',
    help='file path to output excel file'
)

def spot_area(files,output_file):
    """Takes a list of files or a glob compatible path string and iterates over
    all sheets in all files to calculate the area per cell
    
    Args:
        files (list,str): list of file paths or glob compatible file path string
        output_file (str): output excel file path
    
    Raises:
        TypeError: When the files argument is supplied with the wrong data type
    """
    if type(files) == str:
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
        sheets = [i for i in xl.sheet_names if not i.startswith('Sheet')]
        for sheet in sheets:
            df = xl.parse(sheet)
            # get rid of non-ascii characters in column name
            df.rename(columns=lambda x: re.sub(r'\s*\[.*\]',r'',x),inplace=True)
            # omit the blank lines
            df.dropna(how='all',inplace=True)
            # handle different human source pattern by forcing it to match mouse
            df['Source'] = df['Source'].str.replace('Set \d+ \-','')
            # attempt to parse the source field
            source = df['Source'].iloc[0]
            try:
                parsed_source = OrderedDict(source_expr.match(source).groupdict())
            except:
                print(
                    f'WARNING: Unable to extract information from source pattern: {source}'
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
                ("Experiment",experiment),
                ("SheetName",sheet),
                *parsed_source.items(),
                ("Source",df['Source'].iloc[0]),
                ('TotalArea',total_area),
                ('TotalObjects',num_objects),
                ("AreaPerCell",area_per_cell),
            ])
            data.append(expt_result)
    # convert list of OrderedDicts to pandas dataframe
    data = pd.DataFrame(data)
    # export to excel
    if isfile(output_file):
        raise FileExistsError(
            "The output file already exists. Choose a new filename"
        )
    else:
        data.to_excel(output_file,index=False)
    return(data)

if __name__ == "__main__":
    args = parser.parse_args()
    data = spot_area(args.input_path,args.output_file)
    print(
        f'Successfully created output {args.output_file}'
    )