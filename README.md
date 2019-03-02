# java-CSV-to-Exce

## Usage 
java -jar java-csv-to-excel-jar-with-dependencies.jar -t [xlsx|xls] -o [outfile] -d [delimiter] -e [input encoding] -i [infile1:infile2:infile3...]

## Example Usage (Pipe Delimited CSVs)
java -jar java-csv-to-excel-jar-with-dependencies.jar -t xlsx -o myoutfile -d "|" -e UTF-8 -i mysheet1.txt:mysheet2.txt:mysheet3.txt

# Troubleshooting:
1) "Command not found.": If you are using a pipe delimiter it must be escaped with a slash. E.g. "\\|".<br />
2) "(No such file or directory)": Check you have the correct input/output file paths<br />l



## HOWTO build a new executable jar

```
mvn clean compile assembly:single
```