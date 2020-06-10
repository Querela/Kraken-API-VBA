API-Keys
========

Links
-----

- [how to generate a new key & what permission is used for what method](https://support.kraken.com/hc/en-us/articles/360000919966-How-to-generate-an-API-key-pair-)
- [What is a nonce windows?](https://support.kraken.com/hc/en-us/articles/360001148023-What-is-a-Nonce-Window-)
- [What is a nonce?](https://support.kraken.com/hc/en-us/articles/360000906023-What-is-a-nonce-)

Permissions vs. API method
--------------------------

### Funds

| Permission | API method | Example (code) |
| ---------- | ---------- | -------------- |
| Query Funds | [Balance](https://www.kraken.com/features/api#get-account-balance) | [Tabelle10.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle10.cls) |
| | [TradeBalance](https://www.kraken.com/features/api#get-trade-balance) | [Tabelle11.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle11.cls) |
| ? | [TradeVolume](https://www.kraken.com/features/api#get-trade-volume) **???** | [Tabelle20.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle20.cls) |
| Deposit Funds | [DepositMethods](https://www.kraken.com/features/api#deposit-methods) | * |
| | [DepositAddresses](https://www.kraken.com/features/api#deposit-addresses) | * |
| ? | [DepositStatus](https://www.kraken.com/features/api#deposit-status) **???** | * |
| Withdraw Funds | [WithdrawInfo](https://www.kraken.com/features/api#get-withdrawal-info) | |
| | [Withdraw](https://www.kraken.com/features/api#withdraw-funds) | |
| | [WithdrawCancel](https://www.kraken.com/features/api#withdraw-cancel) | |
| ? | [WithdrawStatus](https://www.kraken.com/features/api#withdraw-status) **???** | |
| ? | [WalletTransfer](https://www.kraken.com/features/api#wallet-transfer) **???** | |

### Orders & Trades

| Permission | API method | Example (code) |
| ---------- | ---------- | -------------- |
| Query Open Orders & Trades | [OpenOrders](https://www.kraken.com/features/api#get-open-orders) | [Tabelle12.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle12.cls) |
| | [QueryOrders](https://www.kraken.com/features/api#query-orders-info) | [Tabelle14.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle14.cls) |
| | [OpenPositions](https://www.kraken.com/features/api#get-open-positions) | [Tabelle17.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle17.cls) |
| Query Closed Orders & Trades | [ClosedOrders](https://www.kraken.com/features/api#get-closed-orders) | [Tabelle13.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle13.cls) |
| | [QueryOrders](https://www.kraken.com/features/api#query-orders-info) _(see open orders)_ | [Tabelle14.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle14.cls) |
| | [QueryTrades](https://www.kraken.com/features/api#query-trades-info) | [Tabelle16.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle16.cls) |
| | [TradesHistory](https://www.kraken.com/features/api#get-trades-history) **???** | [Tabelle15.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle15.cls) |
| Modify Orders | [AddOrder](https://www.kraken.com/features/api#add-standard-order) | |
| Cancel/Close Orders | [CancelOrder](https://www.kraken.com/features/api#cancel-open-order) | |

### Ledger

| Permission | API method | Example (code) |
| ---------- | ---------- | -------------- |
| Query Ledger Entries | [Ledgers](https://www.kraken.com/features/api#get-ledgers-info) | [Tabelle18.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle18.cls) |
| | [QueryLedgers](https://www.kraken.com/features/api#query-ledgers) | [Tabelle19.cls](https://github.com/Querela/Kraken-API-VBA/blob/master/example/Tabelle19.cls) |

### Other

| Permission | API method | Example (code) |
| ---------- | ---------- | -------------- |
| Export Data | [AddExport](https://www.kraken.com/features/api#add-history-export) | |
| | [RetrieveExport](https://www.kraken.com/features/api#get-history-export) | |
| | [ExportStatus](https://www.kraken.com/features/api#get-export-statuses) | |
| | [RemoveExport](https://www.kraken.com/features/api#remove-history-export) | |

