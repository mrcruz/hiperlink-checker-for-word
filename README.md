# Hiperlink checker for Word
It goes through your word document and tries to identify any kind of broken URL, testing them and returning URLs that fail the test.

## Development 

python version used 3.7.5

Setup enviroment

    pip install python-docx    
    pip install requests

Build windows executable with 
    pip install pyinstaller
    pyinstaller source/linkTester.py

To test a file
    dist/linkTester/linkTester.exe tests/testDoc.docx