(ns poi-test.excel
  (:gen-class
    :name com.esentri.clojure.excel.reader
    :prefix "cls-"
    :main false
    :methods [[execute [String String] void]])
  (:use [clojure.java.io :only [output-stream]])
  (:import
    (org.apache.poi.xssf.usermodel XSSFWorkbook XSSFRichTextString)
    (org.apache.poi.xssf.streaming SXSSFWorkbook)
    (org.apache.poi.openxml4j.opc OPCPackage PackageAccess)
    (org.apache.poi.xssf.eventusermodel XSSFReader)
    (org.apache.poi.xssf.model SharedStringsTable)
    (org.apache.poi.ooxml.util SAXHelper)
    (org.xml.sax XMLReader InputSource Attributes)
    (org.xml.sax.helpers DefaultHandler)
    ))


(def PRICE 
  {
    :Toner 134.79, 
    :Büroklammern 1.24, 
    :Stift 3.92, 
    :Druckerpapier 4.53, 
    (keyword "Whiteboard Marker") 7.67, 
    :Ordner 1.99})


(defn get-column-as-keyword [^String cellname]
  (keyword (subs cellname 0 1)))

(defn sheet-handler
  [^SharedStringsTable sst rowfn]
  (let [value (atom "")
        cell-info (atom nil)
        row (atom {})
        get-sst-fn (fn [_value] (.getString (.getItemAt sst (Integer/parseInt _value))))
        get-sst (memoize get-sst-fn)
        ]
    (proxy [DefaultHandler]
      []
      (characters [chs start length]
        (swap! value str (String. chs start length)))
      (startElement [^String uri ^String localName ^String name ^Attributes attributes]
        (case name 
          "c" 
            (case (.getValue attributes "t")
              "s" (reset! cell-info [:s (get-column-as-keyword (.getValue attributes "r"))])
              "n" (reset! cell-info [:n (get-column-as-keyword (.getValue attributes "r"))])
              nil)
          "v" (do 
                (reset! value nil))
          nil
        ))
      (endElement [uri localName name]
        (case name
          "v" (let [[type position] @cell-info]
                (case type
                  :s (swap! row assoc position (get-sst @value))
                  :n (swap! row assoc position (Double/parseDouble @value))
                  nil))
          "row" (do (rowfn @row)
                    (reset! row {}))
          nil
          )))))
      


(defn ^XMLReader getParser
  [^SharedStringsTable sst rowfn]
  (let [parser (SAXHelper/newXMLReader)
        _  (.setContentHandler parser (sheet-handler sst rowfn))]
        parser))
(def order-summary-item (atom {}))
(def order-summary-company (atom {}))
(def item-row-count (atom 0))
(def ad0 (fnil + 0 0))
(defn write-cell [out-row index value]
  (.setCellValue (.createCell out-row index) value))
(defn callback [item-sheet row]
;nil)
;  (println row)
  (when (not= "Menge" (:D row))
    (let [item (keyword (:C row))
          amount (:D row)
          price (get PRICE item)
          write-cell (partial write-cell (.createRow item-sheet (swap! item-row-count inc)))
        ]
      (write-cell 0 (:A row)) ;company
      (write-cell 1 (:B row)) ;department
      (write-cell 2 (:C row)) ;item
      (write-cell 3 amount) ;amount
      (write-cell 4 price)
      (write-cell 5 (* price amount))
      (swap! order-summary-company update-in [(:A row) item] ad0 amount)
      (swap! order-summary-item update-in [item] ad0 amount))))

(defn finalize [swb]
  (let [summary-row-count (atom 0)
        summary-sheet (.createSheet swb "Summe Artikel")
        company-row-count (atom 0)
        company-sheet (.createSheet swb "Summe je Firma")]
    (doseq [[item amount] @order-summary-item]
      (let [write-cell (partial write-cell (.createRow summary-sheet (swap! summary-row-count inc)))
            price (get PRICE item)]
        (write-cell 0 (name item))
        (write-cell 1 amount)
        (write-cell 2 price)
        (write-cell 3 (* price amount))))
    (doseq [[company item-amount-map] @order-summary-company]
      (doseq [[item amount] item-amount-map]
        (let [write-cell (partial write-cell (.createRow company-sheet (swap! company-row-count inc)))
              price (get PRICE item)]
          (write-cell 0 company)
          (write-cell 1 (name item))
          (write-cell 2 amount)
          (write-cell 3 price)
          (write-cell 4 (* price amount)))))))

(defn load-wb
  [^String filename ^String out-filename]
  (let [pkg   (OPCPackage/open filename PackageAccess/READ)
        wb    (XSSFWorkbook. pkg)
        swb   (SXSSFWorkbook. 100)
        _     (.setCompressTempFiles swb true)
        item-sheet (.createSheet swb "Einzelposten")
        r     (XSSFReader. pkg)
        sst   (.getSharedStringsTable r)
        parser (getParser sst (partial callback item-sheet))
        iter   (.getSheetsData r)
        sheet-input-source  (.next iter)
        ]; r (.getSheetIndex wb "ESPRIT-Members"))]
    (.parse parser (InputSource. sheet-input-source))
    (.close sheet-input-source)
    (finalize swb)
    (with-open [o (output-stream out-filename)]
      (.write swb o))
    (.dispose swb)
    ))

(defn cls-execute
  [this ^String filename ^String out-filename]
  (load-wb filename out-filename)
  nil)