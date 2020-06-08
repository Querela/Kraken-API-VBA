Examples
========

Contains example Excel file(s) using Macros (`*.xslm`) to demonstrate how to use the **Kraken.com API** VBA modules.

Currently, the [**public API**](https://www.kraken.com/features/api#public-market-data) is completely covered.  
I have exported each sheet modul (`Tabelle1-9`) for easier viewing,
but they are *bound* to the Excel file/structure because of how the input and output is being realised.

The [**private API**](https://www.kraken.com/features/api#private-user-data) is work in progress.  
Currently, the best way to load the Kraken API credentials (key & secret) is using a local `kraken.key` file in the same folder as the Excel file. This is subject to change, and not hard-coded.  
A test routine in `Test.bas` for querying the balance works fine, so the remaining methods are just a matter of time and should work without problems.

The Excel file works as it is (!) and the modules can be used for exploring the code or importing in other Excel files.
(_Probably more likely to import the generic API and Utils modules than the Sheet event handler modules (`Tabelle1-9`) ..._)

TODO
----

- more responsive handlers,
- better error handling (e. g. missing credential file, internet connection loss) and handling of Kraken.com API error messages.
