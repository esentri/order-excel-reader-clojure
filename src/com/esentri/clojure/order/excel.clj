(ns com.esentri.clojure.order.excel
  (:gen-class
    :name com.esentri.clojure.order.excel.reader
    :prefix "cls-"
    :main false
    :methods [^:static [execute [String String] void]])
  (:use [clojure.java.io :only [output-stream]]
        [clojure.set :only [rename-keys]])
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
          "v" (reset! value nil)
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
(def order-summary-company (atom {}))
(def item-row-count (atom 0))
(def ad0 (fnil + 0 0))
(defn write-cell [out-row index value]
  (.setCellValue (.createCell out-row index) value))

(defn callback [item-sheet row]
  (let [order (rename-keys row {:A :company, :B :department, :C :item, :D :amount})]
    (when (not= "Menge" (:amount order))
      (let [item-key (keyword (:item order))
            amount (:amount order)
            price (get PRICE item-key)
            write-cell-fn (partial write-cell (.createRow item-sheet (swap! item-row-count inc)))
          ]
        (write-cell-fn 0 (:company order))
        (write-cell-fn 1 (:department order))
        (write-cell-fn 2 (:item order))
        (write-cell-fn 3 amount) ;amount
        (write-cell-fn 4 price)
        (write-cell-fn 5 (* price amount))
        (swap! order-summary-company update-in [(:company order) item-key] ad0 amount)))))

(defn finalize [swb]
  (let [company-row-count (atom 0)
        company-sheet (.createSheet swb "Summe je Firma")
        write-cell-fn (partial write-cell (.createRow company-sheet (swap! company-row-count inc)))
        ]
    (doseq [[company item-amount-map] @order-summary-company]
      (doseq [[item amount] item-amount-map]
        (let [price (get PRICE item)]
          (write-cell-fn 0 company)
          (write-cell-fn 1 (name item))
          (write-cell-fn 2 amount)
          (write-cell-fn 3 price)
          (write-cell-fn 4 (* price amount)))))))

(defn load-wb
  [^String filename ^String out-filename]
  (let [out-wb   (SXSSFWorkbook. 1)
        item-sheet (.createSheet out-wb "Einzelposten")
        r     (XSSFReader. (OPCPackage/open filename PackageAccess/READ))
        parser (getParser (.getSharedStringsTable r) (partial callback item-sheet))]
    (reset! order-summary-company {})
    (reset! item-row-count 0)
    (.parse parser (InputSource. (.next (.getSheetsData r))))
    (finalize out-wb)
    (with-open [o (output-stream out-filename)]
      (.write out-wb o))
    (.dispose out-wb)
    ))

(defn cls-execute
  [^String filename ^String out-filename]
  (load-wb filename out-filename)
  nil)