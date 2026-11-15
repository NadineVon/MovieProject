# Movie Dashboard
## Table of Content
[Problem Statement](#problem-statement)
[Data Source](#data-source)
[Tools](#tools)
[Dashboard](#dashboard)
[M Code](#m-code)
[Recommendations](#recommendations)
### Problem Statement
Netflix wants to better understand which movie they should produce next, including the most suitable actors and directors. We have a dataset containing movie budgets, box office performance, actors, directors, and genres. Your task is to build an Excel dashboard that provides insights into this dataset. The dashboard should help identify:

- The best-performing actors
- The top movies based on box office metrics
- Director performance
- Genre trends
- Seasonal patterns in movie performance
- Any additional insights that can guide future production decisions
  
The final dashboard should be clear, interactive, and visually compelling, enabling Netflix to make data-driven decisions.
### Data Source
Movie Data : The primary dataset used for this analysis is the "Movie Data Homework.xlsx" file, containing detailed information about each movie's performance (box office and budget), actors, directors and genres. You can download the original datasource here:
You can download the original datasource here:[Movie Dataset Excel file](https://github.com/user-attachments/files/34864174/Movies_Data_Dashboard_Homework_Nadine.xlsx)
### Tools
1. Power Query - I used Power Query for Data Cleaning
2. Excel - I used Excel for Data Analysis
3. Pivot Tables - for Creating the dashboard and Visualizations
### Data Cleaning and Preparation
- Data loading and inspection.
- Handling errors, missing values.
- Data cleaning and formatting. The excel file after the data cleaning & preparation process can be downloaded here: [Movies_Dashboard.xlsx](https://github.com/user-attachments/files/34864421/Movies_Data_Nadine.xlsx)
### Questions for Data Analysis
1. Which top 5 genres were the most successful (Box) these years?
2. Which top 5 actors were the most successful?
3. Top 5 movies by Box and Budget?
4. Seasonality?
### Dashboard 
<img width="793" height="683" alt="image" src="https://github.com/user-attachments/assets/c4ee3955-999a-4696-9dde-231a96246779" />
### M Code


```
let
    Source = Excel.Workbook(File.Contents("C:\Users\nadin\OneDrive\Рабочий стол\Data Analyst\Movies_Data_Homework.xlsx"), null, true),
    #"Movie Data_Sheet" = Source{[Item="Movie Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Movie Data_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Movie Title", type text}, {"Release Date", type date}, {"Wikipedia URL", type text}, {"Genre_First_ID", Int64.Type}, {"Genre_Second_ID", Int64.Type}, {"Director_First_ID", Int64.Type}, {"Cast_First_ID", Int64.Type}, {"Cast_Second_ID", Int64.Type}, {"Cast_Third_ID", Int64.Type}, {"Cast_Fourth_ID", Int64.Type}, {"Cast_Fifth_ID", Int64.Type}, {"Budget ($)", Int64.Type}, {"Box Office Revenue ($)", type number}, {"Column14", type any}, {"Column15", type any}, {"Column16", type any}, {"Column17", type any}, {"Column18", type any}, {"Column19", type any}, {"Column20", type any}, {"Column21", type any}}),
    #"Merged Queries" = Table.NestedJoin(#"Changed Type", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres" = Table.ExpandTableColumn(#"Merged Queries", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Removed Columns" = Table.RemoveColumns(#"Expanded Genres",{"Column15", "Column16", "Column17", "Column18", "Column19", "Column20", "Column21", "Genres.Genre", "Column14"}),
    #"Merged Queries1" = Table.NestedJoin(#"Removed Columns", {"Genre_First_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres1" = Table.ExpandTableColumn(#"Merged Queries1", "Genres", {"Genre"}, {"Genres.Genre"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Expanded Genres1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries2" = Table.NestedJoin(#"Reordered Columns", {"Genre_Second_ID"}, Genres, {"ID"}, "Genres", JoinKind.LeftOuter),
    #"Expanded Genres2" = Table.ExpandTableColumn(#"Merged Queries2", "Genres", {"Genre"}, {"Genres.Genre.1"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Expanded Genres2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres.Genre.1", "Director_First_ID", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns" = Table.RenameColumns(#"Reordered Columns1",{{"Genres.Genre.1", "Genres_second"}}),
    #"Merged Queries3" = Table.NestedJoin(#"Renamed Columns", {"Director_First_ID"}, Directors, {"ID"}, "Directors", JoinKind.LeftOuter),
    #"Expanded Directors" = Table.ExpandTableColumn(#"Merged Queries3", "Directors", {"Director"}, {"Directors.Director"}),
    #"Reordered Columns2" = Table.ReorderColumns(#"Expanded Directors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres_second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries4" = Table.NestedJoin(#"Reordered Columns2", {"Cast_First_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors" = Table.ExpandTableColumn(#"Merged Queries4", "Actors", {"Actor"}, {"Actors.Actor"}),
    #"Reordered Columns3" = Table.ReorderColumns(#"Expanded Actors",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres_second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Merged Queries5" = Table.NestedJoin(#"Reordered Columns3", {"Cast_Second_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors1" = Table.ExpandTableColumn(#"Merged Queries5", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns4" = Table.ReorderColumns(#"Expanded Actors1",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres_second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors.Actor.1", "Cast_Third_ID", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns1" = Table.RenameColumns(#"Reordered Columns4",{{"Actors.Actor.1", "Actors_second"}}),
    #"Merged Queries6" = Table.NestedJoin(#"Renamed Columns1", {"Cast_Third_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors2" = Table.ExpandTableColumn(#"Merged Queries6", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns5" = Table.ReorderColumns(#"Expanded Actors2",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres_second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors_second", "Cast_Third_ID", "Actors.Actor.1", "Cast_Fourth_ID", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns2" = Table.RenameColumns(#"Reordered Columns5",{{"Actors.Actor.1", "Actors_third"}}),
    #"Merged Queries7" = Table.NestedJoin(#"Renamed Columns2", {"Cast_Fourth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors3" = Table.ExpandTableColumn(#"Merged Queries7", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns6" = Table.ReorderColumns(#"Expanded Actors3",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres_second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors_second", "Cast_Third_ID", "Actors_third", "Cast_Fourth_ID", "Actors.Actor.1", "Cast_Fifth_ID", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns3" = Table.RenameColumns(#"Reordered Columns6",{{"Actors.Actor.1", "Actors_fourth"}}),
    #"Merged Queries8" = Table.NestedJoin(#"Renamed Columns3", {"Cast_Fifth_ID"}, Actors, {"ID"}, "Actors", JoinKind.LeftOuter),
    #"Expanded Actors4" = Table.ExpandTableColumn(#"Merged Queries8", "Actors", {"Actor"}, {"Actors.Actor.1"}),
    #"Reordered Columns7" = Table.ReorderColumns(#"Expanded Actors4",{"Movie Title", "Release Date", "Wikipedia URL", "Genre_First_ID", "Genres.Genre", "Genre_Second_ID", "Genres_second", "Director_First_ID", "Directors.Director", "Cast_First_ID", "Actors.Actor", "Cast_Second_ID", "Actors_second", "Cast_Third_ID", "Actors_third", "Cast_Fourth_ID", "Actors_fourth", "Cast_Fifth_ID", "Actors.Actor.1", "Budget ($)", "Box Office Revenue ($)"}),
    #"Renamed Columns4" = Table.RenameColumns(#"Reordered Columns7",{{"Actors.Actor.1", "Actors_fifth"}}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Renamed Columns4",{{"Budget ($)", Currency.Type}, {"Box Office Revenue ($)", Currency.Type}}),
    #"Added Custom" = Table.AddColumn(#"Changed Type1", "ROI", each ([#"Box Office Revenue ($)"]-[#"Budget ($)"])/[#"Budget ($)"]),
    #"Changed Type2" = Table.TransformColumnTypes(#"Added Custom",{{"ROI", Percentage.Type}})
in
    #"Changed Type2"
```
### Recommendations
Top 5 genres are Actions, Comedy, Drama, Sci-Fi, Family. I would recommend Netflix to produce a movie with one of these genres as they brought in more in box office revenue based on the data from 2012 to 2016

<img width="293" height="133" alt="image" src="https://github.com/user-attachments/assets/31a8ca5e-b4c5-4274-91ab-6feedeb1e698" />
