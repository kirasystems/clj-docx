(ns clj-docx.test-helpers
  (:import
   (org.docx4j XmlUtils))
  (:use
   clj-docx.core
   clojure.test))

(defn do-reload []
  (use 'clj-docx.core :reload)
  (use 'clj-docx.core-test :reload))

(defn do-reload-and-run-tests []
  (do-reload)
  (run-tests 'clj-docx.core-test))

(defn tbl-from-file
  [filename]
  (-> filename
      load-wordml-pkg
      extract-body-from-pkg
      extract-tbl-from-body))

(defn tbl-from-default-fixture-file []
  (tbl-from-file "test/fixtures/fixture.docx"))

(defn tbl-rows-from-default-fixture-file []
  (extract-tbl-rows
   (tbl-from-default-fixture-file)))
