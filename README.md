`process_pdfs.py` — первоначальный код неизвестного автора.  
`process_pdfs_1.py`: обрабатываются все pdf файлы в указанной папке, а не только файлы, начинающиеся с "Input".  
 Было:  
 ```
 for pattern in ['Input*.pdf', 'Input?.pdf']
 ```  
стало:  
 ```
 for pattern in ['*.pdf']
 ```  



