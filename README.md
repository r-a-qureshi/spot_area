# spot_area

## Motivation
This project is a command line utility for processing excel files output by the Nikon NIS Elements AR Software. 
The software can be used to detect the number of cells in an image, and measure the area of the region of interest for the cell.
The goal of this project is to provide an easy way to process these excel files and output an excel file with the
average area per cell for each experiment.

## Description
Sheets from input excel files from the Nikon software typically take the form:

| Item   | Source                 | FieldID   | BinaryID            | ND.M   | NumberObjects   | BinaryArea [µm²]   |
|-------:|:-----------------------|----------:|:--------------------|-------:|----------------:|-------------------:|
| 1      | Mus3_011 - UT - UT.nd2 | 1         | Threshold           | 1      | 0               | 0                  |
| 2      | Mus3_011 - UT - UT.nd2 | 2         | Threshold           | 2      | 0               | 0                  |
| …      | …                      | …         | …                   | …      | …               | …                  |
| 1      | Mus3_011 - UT - UT.nd2 | 1         | SpotDetection (GFP) | 1      | 75              | 5931.56            |
| 2      | Mus3_011 - UT - UT.nd2 | 2         | SpotDetection (GFP) | 2      | 71              | 5626.23            |


The goal is to take the sum of the values of BinaryArea across rows where the BinaryID is "Threshold" and divide
by the sum of the values of NumberObjects where BinaryID is "SpotDetection (GFP)." This will calculate the average
area per cell for that particular experimental condition. spot_area will iterate over sheets across multiple files
and calculate the value for each sheet. spot_area will also attempt to parse the Source field to obtain a SubjectID and 
two treatment conditions. If the source cannot be parsed these fields will be filled with "NA." 
spot_area uses the file name as the name of the experiment. spot_area produces an excel file as output that looks like this:

| Experiment   | SheetName     | SubjectID   | Treat1   | Treat2    | Source                    |   TotalArea |   TotalObjects |   AreaPerCell |
|:-------------|:--------------|:------------|:---------|:----------|:--------------------------|------------:|---------------:|--------------:|
| Mus_Data     | Mus3_UT_UT    | Mus3_011    | UT       | UT.nd2    | Mus3_011 - UT - UT.nd2    |     2528    |           2965 |      0.852614 |
| Mus_Data     | Mus3_UT_DMSO  | Mus3_012    | UT       | DMSO.nd2  | Mus3_012 - UT - DMSO.nd2  |     8807.57 |           2875 |      3.0635   |
| Mus_Data     | Mus3_UT_Drug1 | Mus3_013    | UT       | Drug1.nd2 | Mus3_013 - UT - Drug1.nd2 |     1469.79 |           2324 |      0.63244  |
| Mus_Data     | Mus3_UT_Drug2 | Mus3_014    | UT       | Drug2.nd2 | Mus3_014 - UT - Drug2.nd2 |     1137.08 |           3057 |      0.371959 |

## Documentation
### Installation
First make sure that you have the python 3 installed on your windows machine. You can download it 
[here](https://www.python.org/downloads/windows/). When you install python make sure you check the box "Add python to path."

After you have installed python you can install spot_area by opening a command window and typing:
```
pip install https://github.com/r-a-qureshi/spot_area/archive/master.zip
```

### Usage
After you've successfully installed spot_area you will be able to use it from the command line.
Here's how you view the help text.
```
spot_area -h
```
Which will print the following to the command line
```
usage: spot_area [-h] [-i INPUT_PATH] [-o OUTPUT_FILE] [-g] [-r]

Process microscope excel files to calculate area per cell and return an output
excel file. Output excel file should be in a different directory than the
input files.

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT_PATH, --input_path INPUT_PATH
                        File path where excel files are located, can be folder
                        or individual file
  -o OUTPUT_FILE, --output_file OUTPUT_FILE
                        File path to output excel file
  -g, --glob            Use glob to find files specified by input_path
  -r, --recursive       Perform glob search recursively
  ```

  You can use spot_area by providing it with an input path and an output file path. If the input path is a
  folder spot_area will work on all the excel files in that folder and all its subfolders.
  ```
  spot_area -i 'mypath/data_files/' -o 'output_path/my_output_file.xlsx'
  ```

  To run spot_area on a single file instead of a folder you can specify the exact path to the desired file as input
  ```
  spot_area -i 'mypath/data_files/single_file.xlsx' -o 'output_path/single_file_output_file.xlsx'
  ```

  This will create a file in the output_path called my_output_file.xlsx that conforms to the output table above.
  The glob and recursive flags will signal to use glob to find the files specified by the input_path string.

  spot_area can also be used within Python code or a Python interpreter as follows
  ```python
  from spot_area import spot_area
  data = spot_area('mypath/data_files','output_path/my_output_file.xlsx')
  ```
  The data variable will be a pandas DataFrame with the same format as the above output table.