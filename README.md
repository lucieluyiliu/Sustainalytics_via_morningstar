# Sustainalytics_via_morningstar

Sustainalytics_download.py is a python script that cuts the Morningstar universe into batches of 500 securities each, creates excel spreadsheets and calls Morningstar excel add-in to download Sustainalytics risk scores by list-year.<br>

 
Note that Morningstar excel add-in has a built-in error: it cannot correctly align dates when downloading a unbalanced panel for a list of securities. Therefore, the only feasible solutions are downloading time-series security by security or downloading for each list year by year.
