(defproject com.esentri.clojure/order "0.1.2-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "EPL-2.0 OR GPL-2.0-or-later WITH Classpath-exception-2.0"
            :url "https://www.eclipse.org/legal/epl-2.0/"}
  :dependencies [
                  [org.clojure/clojure "1.10.1"]
                  [org.apache.poi/poi "4.1.0"]
                  [org.apache.poi/poi-ooxml "4.1.0"]
                ]
  :main ^:skip-aot com.esentri.clojure.order.core
  :aot [com.esentri.clojure.order.core com.esentri.clojure.order.excel]
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all
                       :jvm-opts ["-Dclojure.compiler.direct-linking=true"]}})