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

(ns xls-to-csv.xls
  (:require [clojure.java.io :as io])
  (:import [org.apache.poi.hssf.usermodel HSSFWorkbook]
           [org.apache.poi.ss.usermodel Row]
           [org.apache.poi.ss.usermodel Cell]))

(defn open-workbook
  "open xls file and return workbook object"
  [f]
  (with-open [s (io/input-stream f)]
    (HSSFWorkbook. s)))

(defn get-sheet
  "get nth sheet form workbook"
  [wb n]
  (.getSheetAt wb n))

(defn get-rows
  "get rows from sheet"
  [s]
  (iterator-seq (.iterator s)))

(defn get-cells
  "get cells from row"
  [r]
  (for [i (range 0 (.getLastCellNum r))]
    (.getCell r i Row/CREATE_NULL_AS_BLANK)))


(defn bad-cell-type-error [c]
  (str "Bad cell type '" (.getCellType c) "'. xls->csv only supports numeric and string
  cells. Cells with formulas are not supported"))

(defmulti get-cell-value
  "get the value of cell based on the cell type"
  #(.getCellType %))

(defmethod get-cell-value Cell/CELL_TYPE_NUMERIC [c]
  (.getNumericCellValue c))

(defmethod get-cell-value Cell/CELL_TYPE_STRING [c]
  (.getStringCellValue c))

(defmethod get-cell-value Cell/CELL_TYPE_BLANK [c]
  "")

(defmethod get-cell-value :default [c]
  (throw (Exception. (bad-cell-type-error c))))

  
(defn whole-number?
  "returns true if n is a whole number"
  [n]
  (or (= 0 (mod n 1)) (= 0.0 (mod n 1.0))))

(defn normalize-numeric-cells [c]
  "convert a whole number float to an integer"
  (if (and (float? c) (whole-number? c))
    (int c)
    c))

(defn fill-rows [rows]
  "fill rows with empty cells so each row has an equal number of cells"
  (let [cols (apply max (map count rows))]
    (map #(concat % (repeat (- cols (count %)) "")) rows)))


(defn workbook-data->vec [in]
  "convert workbook to a multidimensional array of cell data"
  (->> (get-sheet (open-workbook in) 0)
       (map get-cells)
       (map (partial map get-cell-value))
       (map (partial map normalize-numeric-cells))
       (fill-rows)))
