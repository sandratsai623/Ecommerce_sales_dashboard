**Background**
An E-Commerce company sells a variety of products to all over the world.

**Objectives**
The Sales director has asked data analyst to design a sales dashboard that analyzes the sales based on various product categories. The company wants to add user control for product category so that users can select a category and can see the trend month-wise and product-wise accordingly.

**Dataset**
The dataset in file E-Commerce Dashboard dataset.xlsx contains sales data for different product categories.

**Tools**
Excel (Data Analysis Add-in & Combo box)

**Source Code for excel**
A table of Sales and Profit month-wise linked with combo box
=OFFSET($B$2,ROW()-59,$G$48)

A table of Region-wise linked with combo box
=OFFSET($B$25,ROW()-75,$G$48)

Sales 
=SUMIFS('Sales Data'!$H:$H,'Sales Data'!$F:$F,Staging!$G$49)

Quantity 
=SUMIFS('Sales Data'!$I:$I,'Sales Data'!$F:$F,Staging!$G$49)

Profit 
=SUMIFS('Sales Data'!$K:$K,'Sales Data'!$F:$F,Staging!$G$49)
