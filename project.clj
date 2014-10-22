(defproject xls-to-csv "0.1.0"
  :description "A minimal clojure wrapper around Apache POI to convert XSL spreadsheets to CSV."
  :url "https://github.com/kallstrom/xls-to-csv"
  :license {:name "Apache License, Version 2.0"
            :url "http://www.apache.org/licenses/LICENSE-2.0.txt"}
  :dependencies [[org.clojure/clojure "1.6.0"]
                 [org.clojure/data.csv "0.1.2"]
                 [org.apache.poi/poi "3.10.1"]]
  :profiles {:dev {:resource-paths ["test/resources"]}})
