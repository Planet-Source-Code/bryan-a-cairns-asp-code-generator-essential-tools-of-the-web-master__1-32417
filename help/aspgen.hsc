HelpScribble project file.
6
`vRTR-12816Q
0
1
ASP Generator



TRUE

1
BrowseButtons()
0
FALSE
5
10
Scribble10
Welcome




Writing



FALSE
14
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}{\f5\fswiss Arial;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red0\green128\blue0;\red128\green0\blue0;}
\deflang1033\pard\plain\f3\fs32\cf1\b Welcome
\par \plain\f3\fs20\cf0 
\par \plain\f5\fs20\b Welcome To ASP Generator
\par By HTML-Helper\plain\f5\fs20 
\par 
\par To begin, please choose a database by clicking the "Browse\'85" button and selecting the database of your choice.
\par 
\par Please note this program requires you have the MSADO 2.5 Library installed on your computer.
\par If you do not have the MSADO 2.5 library installed, please download it and install it.
\par 
\par This file can be downloaded at: \plain\f2\fs20\cf2\strike http://www.html-helper.com/programs/DAO25.EXE\plain\f2\fs20\cf3 \{link=*! ExecFile("http://www.html-helper.com/programs/DAO25.EXE")\}\plain\f2\fs20\cf0  and it is about 7.5 megs in size.\plain\f3\fs20 
\par }
20
Scribble20
Choose Table and Fields




Writing



FALSE
13
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss\fcharset1 Arial;}{\f4\fswiss\fprq2 System;}{\f5\fswiss Arial;}{\f6\fswiss Arial;}{\f7\froman\fprq2 Times New Roman;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang1033\pard\plain\f3\fs32\cf1\b Choose Table and Fields
\par \plain\f3\fs20\cf0 
\par \plain\f5\fs20 In this section you will need to choose the Table in the Database that you wish to use, and choose the fields that you wish to process.
\par 
\par Choose a Table by selecting the Table Name from the drop down list provided.
\par Once you choose a Table, all the fields in the table are displayed.
\par 
\par To use a field, place a check mark next to it, to ignore a field, uncheck it.
\par 
\par \plain\f6\fs20 Also you will need to select a \ldblquote Link Field\rdblquote . A Link Field is commonly the field you wish to use to denote what recordset you are currently viewing, editing, or deleting. Please choose a Field that has data unique to each record such as a Persons Name, or the Record ID.\plain\f3\fs20 
\par }
30
Scribble30
Options




Writing



FALSE
54
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss Arial;}{\f4\fswiss\fprq2 System;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang1033\pard\plain\f2\fs32\cf1\b Options
\par \plain\f2\fs20\cf0 
\par \plain\f3\fs20 This program allows you to pick and choose what options you want. To choose an option, place a check mark next to it, to ignore an option, un-check it.
\par 
\par Each option will create or alter the pages you are generating.
\par 
\par Also you can edit the names of the pages that are being generated.
\par To edit a page name, click on the option you want to edit, and then enter the new name in the field provided.
\par 
\par In this section you will also need to choose your output directory. This is the directory that all files created will be placed.
\par 
\par The pages created are generated on a template.
\par You can and should alter the templates to meet your desired design.
\par 
\par The page templates are located in the \ldblquote data\rdblquote  sub directory of this program.
\par Each template contains special \ldblquote tags\rdblquote  that tell this program what to do.
\par 
\par Here is a list of all available tags:
\par (Please note all tags must be lowercase and have the prefix and suffix of \ldblquote #\rdblquote )
\par 
\par \plain\f3\fs20\b Page Data tags\plain\f3\fs20 
\par #pagevariables#\tab \tab variables to read from the page post
\par #loadvars#\tab \tab page variables loaded from database
\par #savevars#\tab \tab page variables saved to current recordset
\par #datafile#\tab \tab database filename
\par #includedatabase#\tab include database tag
\par #includesetup#\tab \tab include setup file tag
\par #includemsado#\tab \tab include the MS ADO file
\par #writehead#\tab \tab write header code
\par #writefoot#\tab \tab write footer code
\par #writetable#\tab \tab write table contents code
\par #readtable#\tab \tab read table contents code
\par #navtable#\tab \tab navigation table
\par #tablename#\tab \tab the current table name
\par #add#\tab \tab \tab add record
\par #id#\tab \tab \tab the record id tag
\par 
\par \plain\f3\fs20\b Filename tags\plain\f3\fs20 
\par #pagesetup#\tab \tab setup.asp
\par #pagemsado#\tab \tab msado.asp
\par #pagerecnav#\tab \tab recnav.asp
\par #pagerecdisplay#\tab recdisplay.asp
\par #pagedatabase#\tab \tab database.asp
\par #pagerecadd#\tab \tab recadd.asp
\par #pagerecedit#\tab \tab recedit.asp
\par #pagerecdelete#\tab \tab recdelete.asp
\par 
\par \plain\f3\fs20\b Page Navigation\plain\f3\fs20 
\par #gonav#\tab \tab \tab goes back to Navigation Page (adds "ShowGoBack" from setup.asp)
\par 
\par \plain\f2\fs20 
\par }
40
Scribble40
Page Options




Writing



FALSE
8
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss Arial;}{\f4\fswiss\fprq2 System;}{\f5\froman\fprq2 Times New Roman;}{\f6\fswiss\fcharset1 Arial;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang1033\pard\plain\f2\fs32\cf1\b Page Options
\par \plain\f3\fs20 
\par This section allows you to define the font color, face, and size of all text that it automatically generated by this program.
\par 
\par Please choose the values for each option that you desire.\plain\f6\fs20 
\par }
50
Scribble50
Generate Pages




Writing



FALSE
11
{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss\fcharset1 Arial;}{\f3\fswiss Arial;}{\f4\fswiss\fprq2 System;}{\f5\fswiss\fcharset1 Arial;}{\f6\froman\fprq2 Times New Roman;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\deflang1033\pard\plain\f2\fs32\cf1\b Generate Pages
\par \plain\f2\fs20\cf0 
\par \plain\f3\fs20 This is the final section of this program \endash  the moment of truth.
\par 
\par To begin generating your ASP pages, click the \ldblquote Click here to Finish\rdblquote  button.
\par Your pages will be generated with all the options you have chosen and will be places in the output directory you supplied.
\par 
\par Please note if you selected the option to \ldblquote Overwrite Pages without prompting\rdblquote  you run a risk of overwriting important files that you have already created \endash  so use this option wisely.\plain\f5\fs20 
\par }
1
main="ASP Generator Help",(0,0,511,410),0,,,0
0
0
0
5
*ParagraphTitle
0
Arial
0
11
B...
0
0
0
0
0
0
*PopupLink
0
Arial
0
8
....
0
0
0
0
0
0
*PopupTopicTitle
16711680
Arial
0
10
B...
0
0
0
0
0
0
*TopicText
0
Arial
0
10
....
0
0
0
0
0
0
*TopicTitle
16711680
Arial
0
16
B...
0
0
0
0
0
0
