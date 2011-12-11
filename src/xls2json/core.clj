(ns xls2json.core
  (:require [cheshire.core :as json])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Row Cell))
  (:require [dk.ative.docjure.spreadsheet :as docjure]))

(defmethod docjure/read-cell-value Cell/CELL_TYPE_ERROR [cv _] "ERROR")

(defn read-cell-safe [cell]
  (try (docjure/read-cell cell) (catch Exception e "ERROR")))

(defn col-seq
  [^Row row]
  (docjure/assert-type row Row)
  (let [
        number-of-cells (.getPhysicalNumberOfCells row)
        ]
   (loop [ col-index 0, result []]
      (let [ cell (.getCell row col-index)
             value (read-cell-safe cell)]
        (if (< col-index number-of-cells)
          (recur (+ 1 col-index) (conj result value))
          (conj result value)))))
)

(defn row-seq [^Sheet sheet]
  (docjure/assert-type sheet Sheet)
  (let [wb (docjure/load-workbook "Test.xls")
        sheet (docjure/select-sheet "All" wb)
        rows (docjure/row-seq sheet)
        evaluator (.. wb getCreationHelper createFormulaEvaluator)]
      (map col-seq rows)))

(defn workbook-to-json [^Sheet sheet]
  (docjure/assert-type sheet Sheet)
  (let [rows (row-seq sheet)]
               (json/encode rows)))

(defn spreadsheet-to-json [filename sheet-name]
  (let [wb (docjure/load-workbook filename)
        sheet (docjure/select-sheet sheet-name wb)]
    (workbook-to-json sheet)))

(defn all-spreadsheets-to-json [filename]
  (let [wb (docjure/load-workbook filename)
        sheets (docjure/sheet-seq wb)]
    (json/encode (reduce #(assoc %1 (.getSheetName %2) (workbook-to-json %2)) {} sheets))
  )
)
