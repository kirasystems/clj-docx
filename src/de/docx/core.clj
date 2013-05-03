(ns de.docx.core
  (:import
   (org.docx4j.openpackaging.packages WordprocessingMLPackage)
   (org.docx4j.openpackaging.parts.WordprocessingML MainDocumentPart)
   (org.docx4j.wml Body)
   (org.docx4j XmlUtils)))

(defn load-wordml-pkg [filename]
  "Loads up a WordML Package object from a file"
  (WordprocessingMLPackage/load
   (java.io.File. filename)))

(defn extract-body-from-pkg
  "Extracts Body from a WordML Package"
  [wordml-pkg]
  (-> wordml-pkg
      (.getMainDocumentPart)
      (.getJaxbElement)
      (.getBody)))

;;
;; ASSUMES ONE TABLE IN DOC
;;  --makes assumptions about structure. Bad.
;;
(defn extract-tbl-from-body [body]
  "Extracts Tbl from Body object"
  (-> body
      (.getEGBlockLevelElts)
      first
      (.getValue)))

(defn extract-tbl-rows
  "Returns a collection of JAXB row
   elements from the body element."
  [tbl]
  (.getEGContentRowContent tbl))

(defn row-at [rows idx]
  "Returns TR at index idx in rows"
       (-> rows
           (nth idx)))

(defn row-cells
  "Returns all TCs in a TR"
  [TR]
  (let [cell-content (.getEGContentCellContent TR)]
    (reduce
     #(conj %1 (.getValue %2))
     [] cell-content)))

(defn cells-at-row
  "Directly extracts collection of TCs
   for TR at index idx of Tbl TR ArrayList"
  [rows idx]
  (-> rows
      (row-at idx)
      row-cells))

(defn extract-text-from-R-els [R-els]
  "Extracts and concatenates strings from the
   collection of Text elements from a R element."
  (let [text-els (filter
                  #(= org.docx4j.wml.Text
                      (type (.getValue %1)))
                  R-els)]
    (reduce
     #(str %1 (-> %2 .getValue .getValue))
     "" text-els)))

(defn extract-text-from-R [elem]
  "Extracts string from R element."
  (if (= org.docx4j.wml.R
         (type elem))
    (extract-text-from-R-els
     (.getRunContent elem))))

(defn text-at-cell [cell]
  (reduce
   #(str %1 (extract-text-from-R %2))
   ""
   (-> cell
       .getEGBlockLevelElts
       first ; hmm...always just one?
       .getParagraphContent)))

(defn text-at-cell-idx
  "Given a collection of cells, extracts
   text for the cell at index idx."
  [cells idx]
  (-> cells
      (nth idx)
      (text-at-cell)))

(defn find-cells-with-string
  [cells match-str]
  (let [str-ptrn (re-pattern match-str)]
    (filter
     #(re-find str-ptrn (text-at-cell %1))
     cells)))

(defn has-cell-with-string?
  [cells match-str]
  (seq
   (find-cells-with-string cells match-str)))

(defn find-rows-with-string
  "Finds rows with a cell having text
   matching a particular string"
  [rows match-str]
  (filter
   #(has-cell-with-string? (row-cells %1) match-str)
   rows))

(defn find-first-row-with-string
  "Finds first row with a cell having text
   matching a particular string"
  [rows match-str]
  (first
   (find-rows-with-string rows match-str)))

(defn new-p-with-string [string])
  
