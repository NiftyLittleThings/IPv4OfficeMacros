# IPv4OfficeMacros

Some useful LibreOffice/OO macros for processing/manipulating IPv4 addresses:
- bit shifting with wrap around (currently 32 bit - todo:make generic)
- multi-cell join with arbitrary leading and trailing delimiters (superb, borrowed, see below)
- generate subnets in a string format suitable for POSTing in a query to ElasticSearch
  eg: 1.33.7.0/23 => '( 1.33.7.* OR 1.33.8.* )'


Originally devised to search for firewall, snort and netflow logs in ElasticSearch (created by Graylog).

Note that a separate Python script is used to query and process ES data.


This code is poor/average, borrowed mostly or written in a hurry (and typically late at night) to achieve a
specific task at that point in time.  Then left to bit rot. Time...


Found many of the initial ideas (and the StrJoin code) somewhere else:

StrJoin function - beautifully created by Adam Spiers:
https://stackoverflow.com/questions/1825886/open-office-spreadsheet-calc-concatenate-text-cells-with-delimiters/2417109#2417109

Most useful was this resource:
http://www.pitonyak.org/OOME_3_0.pdf

The above are Objet d'Art compared to this ;)
