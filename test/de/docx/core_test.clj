(ns de.docx.core-test
  (:import
   (org.docx4j.openpackaging.packages WordprocessingMLPackage)
   (org.docx4j XmlUtils)
   (org.docx4j.wml ObjectFactory))
  (:use clojure.test
        de.docx.core
        de.docx.test-helpers))

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

    (testing "Clones Element"
      (let [document-row   (find-first-row-with-string
                            rows "[Document Identifier]")
            cloned-doc-row (clone-el document-row)]
        (is (= (type document-row) (type cloned-doc-row)))
        (is (not= (str document-row) (str cloned-doc-row)))))

    (testing "Sets Text for cell"
      (let [document-row   (find-first-row-with-string
                            rows "[Document Identifier]")
            cloned-doc-row (clone-el document-row)
            first-cell     (first (row-cells cloned-doc-row))]
        (set-cell-text! first-cell "New Text!")
        (is (= (text-at-cell first-cell) "New Text!"))))

    (testing "Sets sets multiple Texts w/line break for cell"
      (let [parties-row   (find-first-row-with-string
                            rows "Parties")
            cloned-doc-row (clone-el parties-row)
            first-cell     (first (row-cells cloned-doc-row))
            set-cell       (set-cell-text! first-cell "New Text!<br>After a line break!")
            cell-content   (.getContent first-cell)
            second-cell-br (-> cell-content
                               second (.getContent) first (.getContent) first)]

        (is (= 3 (count cell-content)))
        (is (= org.docx4j.wml.Br (type second-cell-br)))))

    ;; Again, spike...will be re-written,
    ;; but illustrates basic DSL usage
    (testing "Constructs doc with multiple tables
              and page break from cloned elements"
      (let [tbl                (clone-el tbl)
            document-row       (clone-el
                                (find-first-row-with-string
                                 rows "[Document Identifier]"))
            document-text-cell (first (row-cells document-row))
            party-row          (clone-el
                                (find-first-row-with-string
                                 rows "Parties"))
            party-text-cell-1  (first (row-cells party-row))
            party-text-cell-2  (second (row-cells party-row))
]

        (set-cell-text! document-text-cell "[My AWESOME DOCUMENT]")
        (set-cell-text! party-text-cell-1  "MY PARTIES")
        (set-cell-text! party-text-cell-2  "A list of parties 1) one 2) two 3) three")

        (clear-content! body)
        (clear-content! tbl)
        (add-elem! tbl document-row)
        (add-elem! tbl party-row)
        (add-elem! body tbl)
        (add-elem! body (create-page-br))
        (add-elem! body tbl)

;;        (.save wordml-pkg
;;               (java.io.File. "test/fixtures/save-test.docx"))
))

    ) ;; let
  ) ;; deftest
