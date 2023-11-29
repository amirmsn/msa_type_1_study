This tool is developed by <a href="https://eurdexlab.se/">EurDex Lab</a>.
A Type 1 gage study exclusively evaluates the variability attributed to the gauge. More precisely, it examines the impacts of bias and repeatability on measurements conducted by a single operator using a designated reference part. The tool processes the data and generates a comprehensive report akin to Minitab software, incorporating supplementary analysis outcomes for swift gauge assessment. This tool is developed by Eurdex Lab.
The tool operates based on certain requirements and assumptions, including:

 - Tolerance percentage set at 20%.
 - Sigma level established at 6.
 - Confidence interval maintained at 95%.
 - Criteria for accepting/rejecting Cg and Cgk set at 1.67.
 - Criteria for accepting/rejecting percentage variability stands at 15%.
 - The tool undergoes testing with a sample size of 50.
 - The input file is expected to be in .xlsx format, lacking any header for the record column.

The below image is an example of the output format of the tool.

<img src="https://github.com/amirmsn/msa_type_1_study/blob/main/github_example_01.png">

Prior to the installation of Python, make sure that it is added to the path.
To run the code, go to the folder where it is saved, and in the address bar type cmd to open a command window. Type Python plus the name of the Python code file. Hit Enter to execute it.

The prerequisite libraries and their way of installation before running the code are as follows:
 - pip install pandas
 - pip install matplolib
 - pip install scipy
 - pip install openpyxl
