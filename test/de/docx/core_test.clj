(ns de.docx.core-test
  (:import
   (org.docx4j XmlUtils)
   (org.docx4j.wml ObjectFactory))
  (:use clojure.test
        de.docx.core))

(deftest word-ml

  (let [filename    "test/fixtures/fixture.docx"
        wordml-pkg  (load-wordml-pkg filename)
        body        (extract-body-from-pkg wordml-pkg)
        tbl         (extract-tbl-from-body body)
        rows        (extract-tbl-rows tbl)
        row-1       (row-at rows 1)
        row-1-cells (cells-at-row rows 1)
        first-cell  (first row-1-cells)]

    (testing "Opens WordML docx file, returns WordML Pkg"
      (is (= (type (load-wordml-pkg filename))
             org.docx4j.openpackaging.packages.WordprocessingMLPackage)))

    (testing "Extracts Body element from WordML Pkg"
      (is (= (type (extract-body-from-pkg wordml-pkg))
             org.docx4j.wml.Body)))

    (testing "Extracts Table from Body"
      (is (= (type (extract-tbl-from-body body))
             org.docx4j.wml.Tbl)))

    (testing "Extracts Rows from Body"
      ;; Our fixture has 17 rows.
      (is (= (count (extract-tbl-rows tbl)) 17)))

    (testing "Extracts TCs (table cells) from TR (table row)"
      (is (= org.docx4j.wml.Tr (type row-1)))
      (is (= org.docx4j.wml.Tc (type first-cell)))
      (is (= 2                 (count (row-cells row-1)))))

    (testing "Extracts text from TCs"
      (let [row-1-cells (cells-at-row rows 0)
            row-3-cells (cells-at-row rows 2)
            row-1-cell-1-text (text-at-cell-idx row-1-cells 0)
            row-3-cell-1-text (text-at-cell-idx row-3-cells 0)]
        (is (= "[Document Identifier (e.g., FSInvestmentCORP INVESTMENT MANAGEMENT AGR (aka 371))]"
               row-1-cell-1-text))
        (is (= "Parties" row-3-cell-1-text))))

    (testing "Returns TR by string match"
      (let [matching-row-1 (find-first-row-with-string
                             rows
                             "[Document Identifier (e.g., FSInvestmentCORP INVESTMENT MANAGEMENT AGR (aka 371))]")
            matching-row-2 (find-first-row-with-string
                             rows "Parties")]
        (is (= org.docx4j.wml.Tr (type matching-row-1)))

        ;; These two tests kinda hacky. Hmm.
        (is (= "[Document Identifier (e.g., FSInvestmentCORP INVESTMENT MANAGEMENT AGR (aka 371))]"
               (text-at-cell
                (first (row-cells matching-row-1)))))

        (is (= "Parties"
               (text-at-cell
                (first (row-cells matching-row-2)))))))

    ;;
    ;; THIS IS A SPRINT...WILL BE REWRITTEN
    ;;
    (testing "Updates TC text"
      (let [document-row (.getContent
                          (find-first-row-with-string
                           rows "[Document Identifier]"))

            document-cell (.getContent
                           (first
                            (.getContent
                             (.getValue (first document-row)))))

            docR-copy (XmlUtils/deepCopy (first document-cell))

            document-text  (.getValue
                            (first
                             (.getContent docR-copy)))

            title-row      (.getContent
                            (find-first-row-with-string rows "Title"))
            title-label-cell (.getContent
                              (.getValue (first title-row)))
            title-conts-cell (.getContent
                              (.getValue (second title-row)))

            lp        (.getContent (first title-label-cell))
            cp        (.getContent (first title-conts-cell))
            ]

        (.clear document-cell)
        (.setValue document-text "[THIS IS MY DOCUMENT NAME HERE GUYS]")
        (.add document-cell docR-copy)

        (println document-cell)
        (println (.getValue document-text))

        (.setValue (first (.getContent (first lp))) "Label")
        (.setValue (first (.getContent (first cp))) "Some Contents")

        (.save wordml-pkg (java.io.File. "test/fixtures/save-test.docx"))
        ))

    ) ;; let
  ) ;; deftest
