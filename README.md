# secret-santa
 
# What you need to provide
1. Excel spreadsheet with the collowing columns: Name, Email, Address
2. `no-go.txt` that contains pairs of people that should not be put together. One pair per line in the format `Person1, Person2`

# How to Use
Run the script from the commandline with the following options: `secret-santa.py -p path\to\excel.xlsx -n \path\to\no-go.txt -o \path\to\pickled\matches.matches`

# How to unpickle the matches after the fun is over
Run `thaw-matches.py -m \path\to\pickled\matches.matches`
