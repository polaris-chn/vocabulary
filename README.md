# Vocabulary TXT to XLS
#### I perfer to memorize vocabulary by xls file which support with Microsoft Excel while learning English. I can only, however, find few vocabulary xls files on the network. I built this small project in the period of prepared my GRE test. It helped me and I wish it will be helpful for you.
## Table of Content
* [Background](#40)
* [Requirements](#41)
* [Running](#42)
* [Files Explanation](#43)
<h2 id='40'>Background</h2>
When you want to memorize vocabulary by xls files, you will be frustrating to find there are few vocabulary lists on the network. So, I built this project to convert txt vocabulary files to xls files by python program.
<h2 id='41'>Requirements</h2>
python >= 3.6
<h2 id='42'>Running</h2>

```
python txt2xls.py
python combine_xls_upload.py
```
<h2 id='42'>Files Explanation</h2>
The vocabulary lists folder include some common vocabulary lists, like CET4、CET6、TOEFL and GRE. These are txt files and the program will convert these to xls files. The txt2xls.py file is the program file, you can only run it and you will get three xls files. Those are CET4_edited.xls, CET6_edited.xls and TOEFL.xls. If you want more xls files, you can overwrite small part of program and you will get it.

The newly added file combine_xls_upload.py will generate word_combine.xls which delete the overlap of CET4, CET6 and TOEFL.

