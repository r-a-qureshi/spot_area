from setuptools import setup, find_packages

with open('requirements.txt') as f:
    requirements = f.readlines()

setup(
    name='spot_area',
    version='0.1',
    description=""""Process microscope excel files to calculate area per"""\
        """ cell and return an output excel file. Output excel file should"""\
        """ be in a different directory than the input files.""",
    url='https://github.com/r-a-qureshi/spot_area/',
    author='Rehman Qureshi',
    packages=find_packages(),
    install_requires=requirements,
    entry_points={"console_scripts":["spot_area = spot_area.spot_area:main"]},
    keywords="Nikon microscopy cell spot area"
    zip_safe=False,
)