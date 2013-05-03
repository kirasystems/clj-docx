(ns de.docx.test-helpers
  (:import
   (org.docx4j XmlUtils))
  (:use
   de.docx.core
   clojure.test))

(defn do-reload-and-run-tests []
  (use 'de.docx.core :reload)
  (use 'de.docx.core-test :reload)
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

(defn dump-body-xml [body]
  (XmlUtils/marshaltoString body true true))
