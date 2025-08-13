There have been instances where the partitions could be configured wrongly which will make the WinRE partition fail. This script will detect the issue and remedy it by resizing the OS partition and creating a new WinRE partition. Once done it will copy the files necessary down from the Centrality storage blob into the WinRE partition before activating it.

A fully functional WinRE parition is a requirement for drive encryption to work and be compliant.
