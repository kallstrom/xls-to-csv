;; Copyright 2014 Steven Kallstrom
;;
;; Licensed under the Apache License, Version 2.0 (the "License");
;; you may not use this file except in compliance with the License.
;; You may obtain a copy of the License at
;;
;; http://www.apache.org/licenses/LICENSE-2.0
;;
;; Unless required by applicable law or agreed to in writing, software
;; distributed under the License is distributed on an "AS IS" BASIS,
;; WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
;; See the License for the specific language governing permissions and
;; limitations under the License.

(ns xls-to-csv.core
  (:require [clojure.data.csv :as csv]
            [clojure.java.io :as io]
            [xls-to-csv.xls :refer :all]))

(defn xls->vec [in]
  "converts the first sheet of an xls file to a multidimensional vector"
  (workbook-data->vec in))

(defn xls->csv
  "converts the first sheet of an xls file to csv"
  [in out]
  (with-open [f (io/writer out)]
    (csv/write-csv f (xls->vec in)))
  out)
