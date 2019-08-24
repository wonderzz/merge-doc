# The programming language is python #
## The lib i used include docx、win32com、docxtpl##
## I use so much libs because I don't want Repeated wheels docx is very powerful in create docx,but docxtpl is powerful in edit docx,use them in different demand.
 

# The main steps are： #

## 1.get file in path ##

    def MergeDocx(path):
        """
        find all doc&docx file in path
        """


## 2.swicth doc into docx##

    def ReSaveDoc(path, filename):
    """
    swicth doc into docx
    """


## 3.add file into list ##

    def ReSaveAllDoc(path):
    """
    find file jump to 2
    """


## 4.create merge.docx,combine other file into this。##

    def combine_word_documents(path, files):
    """
    combine file
    """