(ns de.docx.test-helpers
  (:import
   (org.docx4j XmlUtils))
  (:use
   de.docx.core
   clojure.test))

(defn do-reload []
  (use 'de.docx.core :reload)
  (use 'de.docx.core-test :reload))

(defn do-reload-and-run-tests []
  (do-reload)
  (run-tests 'de.docx.core-test))

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

(defn dump-xml [elem]
  (XmlUtils/marshaltoString elem true true))
