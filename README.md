# xls-to-csv

A minimal clojure wrapper around Apache POI to convert XSL spreadsheets to CSV.

### Usage

Add the following to your leiningen dependencies:

```clojure
[xls-to-csv "0.1.0"]
```

### Tutorial

```clojure
(ns example.core
  (:require [xls-to-csv.core :as xls]))

;; convert xls spreadsheet to csv file
(xls->csv "input.xls" "output.csv")

;; or multidimensional vector
(xls->vec "input.xls")
```

XLS spreadsheets containing formulas cannot be converted to CSV, yet. When a cell containing a formula is encountered an error will be thrown.

Please [report any issues](https://github.com/kallstrom/xls-to-csv/issues). Pull requests welcome.

### License

Distributed under the [Apache License, Version 2.0](http://www.apache.org/licenses/LICENSE-2.0.html).
