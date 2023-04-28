# Web Automation for Score Input

This project automates the process of logging into a website, navigating through menus, and inputting scores from an Excel file into a web form. The script uses Selenium for web automation and pandas for data manipulation.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Installation

To run the script, you need to have Python 3.x installed and the following libraries:

1. Selenium: Install it using `pip install selenium`.
2. pandas: Install it using `pip install pandas`.

You also need to download and install the appropriate WebDriver for your browser. For Firefox, download [geckodriver](https://github.com/mozilla/geckodriver/releases) and add its executable to your system PATH.

## Usage

1. Update the following variables in the script with your credentials and the desired values:

```python
user_id.send_keys("YOUR_USER_ID")
user_pwd.send_keys("YOUR_PASSWORD")
```

2. Make sure the Excel file (testExcel.xlsx) is in the same directory as the script.


3. Run the script with `python script_name.py`.

4. The script will open a browser window, log in to the website, navigate through menus, select the appropriate options, and input the scores from the Excel file.

## License
[MIT](https://choosealicense.com/licenses/mit/)