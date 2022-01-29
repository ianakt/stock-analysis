# stock-analysis

Purpose:

To refactor the code provided  to create a faster stock analysis based on trading volume for green energy stocks.

Methods:

Using VBA we used the raw data provided which includes the date, market open, market close,  adjusted close, and volume for 12 stocks. 

The refactored code main difference with original provided is that the refactored code uses arrays to refer to values, instead of the literal cells. 
There was no difference in how the values of starting price, ending price, and total volumes were calculated. However there is a difference in how the
value is referred as in the formula. The refactored code refers to the ticker index, where as the original refers to the cell value. The other major change
was to separate the nested loop in the original code, into two separate loops.

![image](https://user-images.githubusercontent.com/68198233/151661448-51b2dfe5-2baf-41ed-93f0-94221b679581.png)
![image](https://user-images.githubusercontent.com/68198233/151661464-86182998-3373-45e3-bebf-5129f3575917.png)



Analysis

Looking at the performance of the portfolio that Steven added for alternative energy stocks, most went down in 2018 however I couldn't predict the outcome of any particular stock just using a year's worth of data for either volume or return. Please see [here](https://finance.yahoo.com/chart/DQ#eyJpbnRlcnZhbCI6IndlZWsiLCJwZXJpb2RpY2l0eSI6MSwidGltZVVuaXQiOm51bGwsImNhbmRsZVdpZHRoIjozLjU2MzIxODM5MDgwNDU5NzYsImZsaXBwZWQiOmZhbHNlLCJ2b2x1bWVVbmRlcmxheSI6dHJ1ZSwiYWRqIjp0cnVlLCJjcm9zc2hhaXIiOnRydWUsImNoYXJ0VHlwZSI6ImxpbmUiLCJleHRlbmRlZCI6ZmFsc2UsIm1hcmtldFNlc3Npb25zIjp7fSwiYWdncmVnYXRpb25UeXBlIjoib2hsYyIsImNoYXJ0U2NhbGUiOiJsaW5lYXIiLCJwYW5lbHMiOnsiY2hhcnQiOnsicGVyY2VudCI6MSwiZGlzcGxheSI6IkRRIiwiY2hhcnROYW1lIjoiY2hhcnQiLCJpbmRleCI6MCwieUF4aXMiOnsibmFtZSI6ImNoYXJ0IiwicG9zaXRpb24iOm51bGx9LCJ5YXhpc0xIUyI6W10sInlheGlzUkhTIjpbImNoYXJ0Iiwi4oCMdm9sIHVuZHLigIwiXX19LCJzZXRTcGFuIjp7Im11bHRpcGxpZXIiOjUsImJhc2UiOiJ5ZWFyIiwicGVyaW9kaWNpdHkiOnsicGVyaW9kIjoxLCJpbnRlcnZhbCI6IndlZWsifSwibWFpbnRhaW5QZXJpb2RpY2l0eSI6dHJ1ZSwiZm9yY2VMb2FkIjp0cnVlfSwibGluZVdpZHRoIjoyLCJzdHJpcGVkQmFja2dyb3VuZCI6dHJ1ZSwiZXZlbnRzIjp0cnVlLCJjb2xvciI6IiMwMDgxZjIiLCJzdHJpcGVkQmFja2dyb3VkIjp0cnVlLCJldmVudE1hcCI6eyJjb3Jwb3JhdGUiOnsiZGl2cyI6dHJ1ZSwic3BsaXRzIjp0cnVlfSwic2lnRGV2Ijp7fX0sImN1c3RvbVJhbmdlIjpudWxsLCJzeW1ib2xzIjpbeyJzeW1ib2wiOiJEUSIsInN5bWJvbE9iamVjdCI6eyJzeW1ib2wiOiJEUSIsInF1b3RlVHlwZSI6IkVRVUlUWSIsImV4Y2hhbmdlVGltZVpvbmUiOiJBbWVyaWNhL05ld19Zb3JrIn0sInBlcmlvZGljaXR5IjoxLCJpbnRlcnZhbCI6IndlZWsiLCJ0aW1lVW5pdCI6bnVsbCwic2V0U3BhbiI6eyJtdWx0aXBsaWVyIjo1LCJiYXNlIjoieWVhciIsInBlcmlvZGljaXR5Ijp7InBlcmlvZCI6MSwiaW50ZXJ2YWwiOiJ3ZWVrIn0sIm1haW50YWluUGVyaW9kaWNpdHkiOnRydWUsImZvcmNlTG9hZCI6dHJ1ZX19XSwic3R1ZGllcyI6eyLigIx2b2wgdW5kcuKAjCI6eyJ0eXBlIjoidm9sIHVuZHIiLCJpbnB1dHMiOnsiaWQiOiLigIx2b2wgdW5kcuKAjCIsImRpc3BsYXkiOiLigIx2b2wgdW5kcuKAjCJ9LCJvdXRwdXRzIjp7IlVwIFZvbHVtZSI6IiMwMGIwNjEiLCJEb3duIFZvbHVtZSI6IiNmZjMzM2EifSwicGFuZWwiOiJjaGFydCIsInBhcmFtZXRlcnMiOnsid2lkdGhGYWN0b3IiOjAuNDUsImNoYXJ0TmFtZSI6ImNoYXJ0IiwicGFuZWxOYW1lIjoiY2hhcnQifX19fQ--) Though DQ was not impressive in 2017 or 2018 it grew over time. I do believe that the reasoning is sound in choosing alternative energy stocks, only time will tell if these investments are worth while, however developments in the car industry as well as a culture change to renewables seems promising.

Conclusion: 

Our refactoring had no effect on the stock performance, and produced much faster results because arrays were used for the starting prices, ending prices, and total volume. An additional benefit is that the refactored code is very similar to the original in organization. ![image](https://user-images.githubusercontent.com/68198233/147840544-f5d6ddbc-e953-45fb-9e17-e70e685f8a13.png)

