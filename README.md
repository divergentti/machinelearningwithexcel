# Virtual Doctor with Excel
A few macros for the Excel. Not excaclty machine learning, but close.

MOTIVATION and Goal

Machine Learning in general is very interesting topic and I wanted to give my best in given time frame. 

My plan was to download diseases-symptoms database, either directly or by scraping, clean the data and make an application with Octave. Most important is to get data in, then focus to other phases of the project. For the ML standpoint, I need to have data which can be classified. Most likely the csv-format will be best for Octave.  Data should be cleaned with Excel. Cleaning should be done so that outcome classification be produced from the data. Overall plan would be as follows:
 
When data is in the csv-format, I need to learn how to use Octave and make asked two or more models by using Octave code. My plan is to look from Google examples and use existing samples as basis of my code. 

Compared to Java or VBA language in Octave does not look so complex, but documentation seems to be like in the Unix in general, not very descriptive or intuitive.

DISCUSSION

Project description was at the beginning pointing to Mayo Clinic web site, but it was not clear shall project cover all listed diseases and symptoms found from the Mayo Clinic.  It was changed some day to Medlineplus.

I planned to try to split symptoms to their own words and then create vectors from symptom to disease. I guessed best way is to split to 3D array/vector, where depth Z contains split words from symptoms. I believed, that they can be used for weighting the match and for regressions and other ML algorithms.

First issue was that the Octave cannot do web scraping, therefore I had to study which tools works best in this case. Then I ended up to Web scraping issues:

-	There are many web scraping utilities available, commercial and free applications. I decided to test out htttrack . It took about 18 hours to download Mayo Clinic website to the PC over 100 Mbit/s connection, but those files are having unique html-file names and therefore it is complex to make a script, which organizes them so that they can be stripped out from the html-code and imported to the Excel. Data size was 1,97 gigabytes holding 61 614 files (just html, no pictures or videos etc media). I consider this as a show stopper for the httrack: not able to categorize nor import data to Octave or to Excel. I had to look for other tools.
-	I found a Web Scraper  plugin for the Chrome shall be easiest and free tool. This point I found that the project description had changed pointing to Medlineplus, not Mayo Clinic. 
-	Try #1: problem with multi lines (list items) items are scraped so that they are all together without separator. I found multi-item working somewhat promising, but for some reason full scale scrape gave nothing. Test run took 20 hours. Then I did again small scale testing and finally got promising data. However, full scale scraping failed again. 
1.	I decided to look other tools for scraping. Biggest issue with Web Scraper was that regexr (regular expression) can’t replace list separators (html <li>) into other characters and all items is scraped into on long word. Therefore, I tried following tools:
a.	Data Miner – need to pay and complex.
b.	iRobot: too complex and very old.

Spent a lot of time with these tools, but they are either complex to use or commercial and rather expensive. 

Finally, I got an idea to go back to Web Scraper plugin and scrape symptoms etc. in a html format and then later in Excel clean out html codes and replace list item (balloons) with a dot (separation character).

PC to scraped again some 20 hours. This method worked fine. Importing csv into Excel works like a charm and found 4396 diseases.  Data looked promising overall. I cleaned the data by:
-	replaced </li> with “.” -> overcomes the list problem. Now each line ends to “.” which could be used as a tag for symptom etc. separation later in Octave or VBA code.
-	removed html codes with replacing <*> to nothing.
-	Excel diseases-db.xslx size is 4084 (4 megs) and in csv format size is 10143 (10 megs).

Phase 1 was completed fine. Data is usable.

Then I run into show stopper with Octave. I needed to make an 3D Vector/Array for the Octave, where we later load the csv file. Found, that Octave do not support 3-dimensional (z,y,z) Vectors/Arrays at all.  I planned to make two separate Vectors and build some sort of relationship for split words.
Next I hit into problems with Octave version. In the lecture videos, download link pointed to version 3.2.4, which is very old and IO is very limited. I uninstalled old and installed version 4.2, but for some reason pkg load io behaved strange in my primary PC. I installed Octave into another PC, and it seemed to work. Decided to clean original PC installation, and after reinstallation the IO module worked. However, importing csv file with strings remained still big issue. After testing many methods, I got this method working for import: 

[DisName,DisHref,Summary,Causes,Symptoms,ExmTests,Treatment,Outlook,Poscomp,AltNames] = textread('disdb.csv','%s %s %s %s %s %s %s %s %s %s','delimeter',',',1) …. 

Resulting vector was ugly, not usable at all. I googled around and realized, that %s should be replaced with %q, but “%q” option (shall read whole text row) do not work with textread! This means, that importing with textread is out of question. I tried Octave’s xlsread from xlsx files. This is not working either for strings. I tried saving csv as a text etc. Googling around for a few days. I did not find any help nor examples. Felt pressure to deliver project and considered this as a show stopper for full scale Octave project.

Main goal was to make a Virtual Doctor. I needed rethink with which tool I can categorize strings and import them into Octave and then deliver project in given timeframe. Started building Excel with export sheet OctaveLink, but unfortunately, Excel do not have Octave API. Therefore, I had to complete whole project with Excel VBA and understand, that I have to do this in a “poor man’s way” with those constraints which Excel has.

My new plan was to build a word list of symptoms and then find full and partial matches in relation to disease for classification and pass classification to SVM etc. modules. Theoretical point of view, I end up to SVM Classification method. A linear SVM requires solving a quadratic program with several linear constraints. Linear classification is possible to do with Excel, but classifier margins I was not able to resolve.  

I also investigated what is ML and Excel situation now and it seems that Microsoft bought Revolution Analytics (R-focuses) and ML is in fact available at Azure and DataScope. I do have an Azure account, but sharing Azure API key with peers reviewing my work is not tempting. However, it maybe so that ML capabilities will be added to future Excel-versions and therefore this exercise should be fine.
