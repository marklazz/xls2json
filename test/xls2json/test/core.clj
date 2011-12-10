(ns xls2json.test.core
  (:require [cheshire.core :as json])
  (:require [dk.ative.docjure.spreadsheet :as docjure])
  (:use [xls2json.core] :reload)
  (:use [clojure.test]))

(def cell-33-3
  (let [wb (docjure/load-workbook "Test.xls")
        sheet (docjure/select-sheet "All" wb)
        evaluator (.. wb getCreationHelper createFormulaEvaluator)]
    (read-cell-safe (.getCell (.getRow sheet 33) 3))))

(deftest value-in-sheet-is-read-properly
  (is (= 0.0727 cell-33-3)))

(deftest value-in-sheet-is-same-as-json
  (is (= 0.0727 (get (get (json/decode (spreadsheet-to-json "Test.xls" "All")) 33) 3))))

(deftest value-in-sheet-seq-is-same-as-selecting-sheet-individually
  (is (= 0.0727 (get (get (json/decode ((json/decode (all-spreadsheets-to-json "Test.xls")) "All")) 33) 3))))
