### How to use this?

This script help in spliting large excel sheet into multiple smaller excel sheet with specified number of rows

```go build .```

### Usage 
```./vendoruploadsplit -help```

To add generic_creditcard to vendors 

```./vendoruploadsplit -filename vendors.xlsx -payment-method generic_creditcard```


To remove generic_creditcard from vendors 

```./vendoruploadsplit -filename vendors.xlsx -payment-method generic_creditcard operation remove```


To specify 5000 rows in each file 

```./vendoruploadsplit -filename vendors.xlsx -payment-method generic_creditcard -operation remove -rowsPersheet 5000```
