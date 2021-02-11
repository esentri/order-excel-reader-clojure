(ns com.esentri.clojure.order.core
  (:gen-class)
  (:require
    [com.esentri.clojure.order.excel :as excel])
  )




(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  (println "Start")
  (time (excel/load-wb "order_overview.xlsx" "order_output.xlsx"))
  (println "Finish"))
