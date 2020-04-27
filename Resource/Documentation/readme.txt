System Requirement: 
1. JDK (or JRE) 7.


HOW TO RUN

1. Download the software (.zip) from the given link
2. Extract the .zip file. Keep the folder structure intact.
3. Start the software by double clicking SNA.jar


Input Excel file Criteria: 

1. Must be of .xls (Excel 97-2003) format.
2. Input data must appear in first three columns. (A,B,C)
3. Only the rows that contain Number in each of the column A,B,C are considered.


Further improvements:

1. Gephi provides a community detection algorithm and modularity. Then can be incorporated.
2. Gephi provides some automatic layouting algorithm to neatly laying out the graph. They can be incorporated.
3. Introduce option to control Node size scaling, edge color transformation based on centrality measures.
4. Sometimes the graph is not properly shown until a mouse click is given on the graph area. Also the graph fails to render in rare cases. These bugs are more likely to occur from gephi library. Trying to remove those.
5. Though the calculated data seems to be okay, I will generate similar graph in ORA and match the result with that of the software.


Platform info:

Developed using Netbeans IDE 7.2 in Java platform in Windows.
Library used: gephi toolkit (https://gephi.org/toolkit/) for graph visualizaton and centrality measurement.
jxl library for reading/writing excel file.

