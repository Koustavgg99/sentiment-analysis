Approach Explanation
1. Load Necessary Libraries:
- Various Python libraries are used for different tasks:
- `nltk` for natural language processing, including tokenization and stop words.
- `re` for regular expressions to identify specific patterns in text.
- `textstat` for text statistics such as syllable count.
- `openpyxl` for reading from and writing to Excel files.
- `os` for file and directory operations.
2. Download NLTK Resources:
- Ensure the necessary NLTK resources are downloaded. These resources include
tokenizers and stop words essential for processing the text.
3. Load Stop Words and Dictionaries:
- Custom stop words are loaded from the `StopWords` directory.
- Positive and negative word dictionaries are loaded from the `MasterDictionary` directory.
4. Text Cleaning and Tokenization:
- The script tokenize the text into words, converts them to lowercase, and removes stop
words and non-alphabetic characters.
5. Calculate Sentiment Scores:
- Sentiment scores are calculated based on the presence of words in positive and
negative dictionaries. Polarity and subjectivity scores are derived from these values.
6. Readability Analysis:
- Various readability metrics are computed, including average sentence length,
percentage of complex words, Fog Index, and average number of words per sentence.
7. Other Metrics:
- Additional metrics include the count of complex words, total word count, syllable count
per word, personal pronoun count, and average word length.
8. Update Excel File:
- The script updates an existing Excel file (`Output.xlsx`) with the calculated metrics for
each text file found in the `extracted_articles` directory.
Dependencies Required
- Python Libraries:
- `nltk`: For natural language processing tasks.
- `textstat`: For calculating text statistics like syllable count.
- `openpyxl`: For working with Excel files.
- `re`: Part of the standard library for regular expressions.
- `os`: Part of the standard library for file and directory operations.
You can install the required libraries using the following commands:
************************************************
pip install nltk textstat openpyxl
**********************************************
How to Run the Script
1. Ensure All Dependencies are Installed:
- Install the required Python libraries using the commands provided above.
2. Prepare the Directory Structure:
- Your directory should include:
- A `StopWords` directory with all the stop words files.
- A `MasterDictionary` directory with `positive-words.txt` and `negative-words.txt`.
- An `extracted_articles` directory with the text files you want to analyze.
- An `output.xlsx` file with the appropriate column names.
3. Download NLTK Resources:
- Run the following code snippet to download the necessary NLTK resources:
***********************************
import nltk
nltk.download('punkt')
nltk.download('stopwords')
*************************************
4. Save and Run the Script:
- Save the provided script as `sentiment_analysis.py` in your working directory.
- Execute the script using the following command:
***********************************
python sentiment.py
*************************************
This script processes all text files in the `extracted_articles` directory and updates the
`output.xlsx` file with the calculated metrics for each file. Each file's results are added to the
row corresponding to its `URL_ID` in the Excel sheet.
