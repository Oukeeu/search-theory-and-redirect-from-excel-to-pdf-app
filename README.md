# search-theory-and-redirect-from-excel-to-pdf-app
Desktop application to automate theory search from excel and redirect to theory pdf page on demand

This paper is a term paper in the first year of study at St. Petersburg University.

Application for easy search of information on specified parameters from excel document and redirect to pdf page of reference book, which contains the theory of search, which implements the following logic:
1) The user enters a query in the application interface - in this case, a ship part (the application was developed for a shipbuilding company).
2) The search for the specified element, which is entered in the search engine, is performed from the prepared knowledge base in the form of an excel document.
3) Brief information is displayed on the screen: part type, serial number, short description, dimensions, etc.
4) You are prompted to go to the pdf page where the part is described.
5) Click on the button to open the pdf document on the page describing the part.   

Stack:
1) Python
2) PyQt5
3) Excel 
4) openpyxl
5) PyFPDF

