# checkmyxl
Automate validation of your data in Microsoft Excel sheet.  
Built on top of [xlwings]([https://www.xlwings.org).

## Installation (macOS)
### With Conda package manager
See [Conda documentation](https://docs.conda.io) for more details.
```
# 1. Clone this repository
git clone https://github.com/mateuszrezler/checkmyxl.git

# 2. Go to `checkmyxl` directory
cd checkmyxl

# 3. Create `checkmyxl` environment with dependencies
conda create -c conda-forge -n checkmyxl appscript pandas psutil xlwings

# 4. Activate new environment
conda activate checkmyxl

# 5. Install xlwings addin
xlwings addin install

# 6. Create a new project
# If `-s` argument was passed, a sample data is generated
# (recommended for getting started)
python startproject.py -s
```

## Test run
For a new project with sample data generated.
```
# Just run `book.py` file
python book.py

# or select from Excel's menu: xlwings > Run main
```

