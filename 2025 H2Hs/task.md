For each of the rows in the "Shoots" sheet of "2025 H2H Shoots.xlsx" where the "imported" column equals 0, go to the URL given by the URL column. Scrape the HTML table present at this URL using the functions given in "main.py", returning a table with the below columns:
-	Rank (called “Pos.” on the URL)
-	Name (called “Athlete” on the URL)
-	Club (called “country or state code” on the URL. Remove the numeric code and -, such that "731 - Meriden Archery Club" becomes "Meriden Archery Club" and simlar)
-	Score (called “Tot.” on the URL)
Compile the results from all the URLs into new sheets of "2025 H2H Shoots.xlsx", with each sheet called the “Event Name” given by the below dataset. Several event sheets are already compelete, use these as example formats. Change the "Imported" column to 1 when each shoot is imported.