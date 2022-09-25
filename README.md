# DNT_cabin_availability_system
This is some simple code to scrape visbook's system to get an overview of the different cabins availability.

The system was created as I was WAY to irritated over finding an available cabin. 2 lunch breaks later, a saturday morning used to learn formating with openpyxl and vóla. A crappy python script that does the job. somewhat.


<h3>Controll document:</h3>

* **H#** indicates a heading in the final document
* **"STOP"** Stops the system from checking cabins further down the list

This was done in 4-5 hours.

<h3>Updates</h3>
<h4>2022-09-25</h4>

* Tried to look at problem with ConnectionAbortedError: [WinError 10053] An established connection was aborted by the software in your host machine. This turnes out to be a problem with the time the request for updating the values takes. Not solved. Probably some stuff that can be done. Will post as an issue.

* Some small bug fixes with dates.
