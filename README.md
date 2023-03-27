The "Synonyms" program is a Python script that allows for finding synonyms and changing words in a .docx file. It utilizes modules such as BeautifulSoup, requests, docx, collections, prettytable, and re.

The "Synonyms" class has three methods: "change_all", "change_my_choice", and "show_statistics". The "change_all" method changes all words in the file, while "change_my_choice" allows for selecting specific words to be changed. The "show_statistics" method presents statistics on the frequency of individual words in the .docx document.

The program retrieves synonyms from the "synonimy.pl" website and uses them to change words in the document. Words for which no synonyms are found are marked with an asterisk, and words for which there are multiple synonyms are marked with a dash. The program's output can be saved in a .docx file.

The program is user-friendly thanks to the use of colors to highlight individual words in the statistics and to distinguish changed words in the resulting .docx file.
