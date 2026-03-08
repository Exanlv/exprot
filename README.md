## Exprot

### NOTE: WIP. NOT RECOMMENDED FOR PRODUCTION USE

Low memory usage Excel export creation

Client has way too much data, but insists on an xlsx export rather than easy to chunk CSV? Exprot is your friend.

Individual batches of data are written to disk, rather than loading _all_ data into memory at once, making it possible to write massive datasets without going over memory limits. A test data set of 200.000 rows used only ~4MB of memory.
