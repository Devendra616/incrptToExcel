# Description

INCRPT file is a standard report that is generated from our incentive software.
This application reads the INCRPT file and pushes the required data to excel.

## Run

Use node command to start the application

```bash
node index.js
```
or simply the npm command

```
npm start
```


## Sample Input

The sample input is required data with some headers and other un-used information and summary.

```
16 B2081 S NAMESWAR RAO     P05  06       HEM.OPTR.I  26.0 1545.00  3410.00    0.00 231.75    0.00     5186.75   888.00  133.20    0.00  1021.20     0.00     0.00

 17 B1442 VIJAI KUMAR        P06  09   EM OPTR. Gr-II  26.0 2317.50  6227.00    0.00 347.62    0.00     8892.12  1110.00  166.50    0.00  1276.50     0.00     0.00

 18 B2098 SUBHAMOY ROY       P15  06       HEM.OPTR.I  26.0 1545.00  5260.00    0.00 231.75    0.00     7036.75   888.00  133.20    0.00  1021.20     0.00     0.00

 19 B0931 AITU RAM           P23  10    M.HEM.OPR(D)I  25.0 2475.96  3569.00    0.00   0.00    0.00     6044.96  1067.31    0.00    0.00  1067.31     0.00     0.00

 20 B1059 SOHAN SINGH        P23  10    M.HEM.OPR.GRI  25.0 2475.96  3527.00    0.00   0.00    0.00     6002.96  1067.31    0.00    0.00  1067.31     0.00     0.00

```

## Sample Output
![sample output](https://res.cloudinary.com/nmdc/image/upload/v1621937882/github%20readme/incrpt.png)