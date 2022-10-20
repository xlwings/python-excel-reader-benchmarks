# Python Excel Reader Benchmarks

1. Clone this repo
2. Install an xlwings trial license key by running the following on a Terminal/Command Prompt/Anaconda Prompt (get one at https://www.xlwings.org/trial):

   `xlwings license update -k my-key`
3. Install the dependencies into a virtual env or Conda env by running `pip install -r requirements.txt`
4. Run the benchmark like so: `python benchmark.py`. To export the report to a file, run: `python benchmark.py > myreport.txt`.

You can edit/add test cases by editing the `test_cases` dictionary at the bottom of `benchmark.py`.

Reference: https://towardsdatascience.com/benchmarking-python-code-with-timeit-80827e131e48