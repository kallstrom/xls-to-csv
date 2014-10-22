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

(ns xls-to-csv.core-test
  (:require [clojure.test :refer :all]
            [clojure.data.csv :as csv]
            [clojure.java.io :as io]
            [xls-to-csv.core :refer :all])
  (:import [java.io File]))

(defn read-csv [csv]
  (csv/read-csv (slurp csv)))

(defn test-xls->csv [f]
  "convert a test file to csv and compare to expected output"
  (let [in       (io/resource (str "in/" f ".xls"))
        temp     (File/createTempFile f ".csv")
        out      (read-csv (xls->csv in temp))
        expected (read-csv (io/resource (str "expected/" f ".csv")))]
    (.delete temp)
    (is (= out expected))))

(deftest simple-output
  (test-xls->csv "simple"))

(deftest formula-output
  (is (thrown? Exception (test-xls->csv "formula"))))

(deftest precision-output
  (test-xls->csv "precision"))

(deftest missing-cells-output
  (test-xls->csv "missing-cells"))

(deftest staggered-output
  (test-xls->csv "staggered"))
