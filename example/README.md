Examples
========

Contains example Excel file(s) using Macros (`*.xslm`) to demonstrate how to use the **Kraken.com API** VBA modules.

Currently, the [**public API**](https://www.kraken.com/features/api#public-market-data) is completely covered.  
I have exported each sheet modul (`Tabelle1-9`) for easier viewing,
but they are *bound* to the Excel file/structure because of how the input and output is being realised.

The [**private API**](https://www.kraken.com/features/api#private-user-data) is work in progress.  
Currently, the best way to load the Kraken API credentials (key & secret) is using a local `kraken.key` file in the same folder as the Excel file. This is subject to change, and not hard-coded.  

- [Private user data](https://www.kraken.com/features/api#private-user-data) is more or less completely covered (see `Tabelle10-20`). I intentionally excluded examples for exports.
- [Private user trading](https://www.kraken.com/features/api#private-user-trading) will not be included.
- [Private user funding](https://www.kraken.com/features/api#private-user-funding) may come later. Something like deposit history etc. seems useful.

My examples are only intended to query (private user) information, so I will not include routines/examples etc. that may modify anything.
So only [viewing API/Key permissions](https://support.kraken.com/hc/en-us/articles/360000919966-How-to-generate-an-API-key-pair-) should suffice.
You may want to compare the API docs (methods names) and the key permissions, to see what you really need.  
Note, that an used API key pair may not work because of different `nonce` values,
(e. g. the Kraken Pro App may use a finer resolution, in my Excel sheet I use milliseconds)
so creating a new key pair just for VBA is required(?) and 'best practice', as the key pair can later easily be disabled.

The Excel file works as it is (!) and the modules can be used for exploring the code or importing in other Excel files.
(_Probably more likely to import the generic API and Utils modules than the Sheet event handler modules (`Tabelle1-9`) ..._)

TODO
----

- more responsive handlers,
- better error handling (e. g. missing credential file, internet connection loss) and handling of Kraken.com API error messages.
- refactoring the API routines into an class for clearer access, using classes for return values (also easier/clearer access?)
