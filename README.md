Kraken-API-VBA
==============

Kraken.com API using Visual Basic for Applications (VBA, in Excel)

The following files (_modules_) are more or less linked together and should be imported into and Excel file together to be able to work.

The utility modules:

- `WebUtils.bas` -- for parsing/printing JSON, HTTP GET/POST requests
- `CryptoUtils.bas` -- for Hashing, Signing, Byte conversion etc.
- `ExcelUtils.bas` -- with some helper functions for Excel

The public interface (to be used from other modules etc.):

- `API.bas` -- public interface for Kraken.com API

Test code:

- `Test.bas` -- with some test routines

Example Use-Case(s):

- `example/` folder


Copyright and License Information
---------------------------------

Copyright (c) 2020 Querela.  All rights reserved.

See the file "LICENSE" for information on the history of this software, terms &
conditions for usage, and a DISCLAIMER OF ALL WARRANTIES.

All trademarks referenced herein are property of their respective holders.

**NOTE**: Some code has been used (as is or adapted) from online sources, like [StackOverflow](https://stackoverflow.com/) or blogs.
I tried to include the link to the original code on top of the file or in the function itself but I may have not been consistent with this, _so all links (and some more) have been listed in the `API.bas` file at the end._  
Credits also to [krakenex](https://github.com/veox/python3-krakenex) and the Kraken.com API Docs.

Please contact me if there are license conflicts or something related... and I will try to correct as much as I can.

