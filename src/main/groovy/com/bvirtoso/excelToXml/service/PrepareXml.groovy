package com.bvirtoso.excelToXml.service

import groovy.xml.MarkupBuilder
import org.apache.commons.logging.Log
import org.apache.commons.logging.LogFactory
import org.springframework.stereotype.Component

@Component
class PrepareXml {
    protected  final static Log LOG = LogFactory.getLog(PrepareXml.class)

    private final ExcelFileService excelFileService;

    PrepareXml(final ExcelFileService excelFileService){
        this.excelFileService = excelFileService
    }
    void prepareXmlFromExcel(String fromexcelFile, String toXmlFile) {

        def xmlInput = excelFileService.readExcel(fromexcelFile, 0)

        def xmlWriter = new StringWriter()
        def xmlMarkup = new MarkupBuilder(xmlWriter)
        xmlMarkup.setDoubleQuotes(true)
        xmlMarkup.setOmitEmptyAttributes(true)
        xmlMarkup.'Persons'('xmlns:x': 'http://www.groovy-lang.org') {
            xmlInput.each {
                node ->
                    def parentKeySet = node.keySet()
                    xmlMarkup.'Person'((parentKeySet.getAt(0).toString()): node[parentKeySet.getAt(0)],
                            (parentKeySet.getAt(1).toString()): node[parentKeySet.getAt(1)],
                    ) {
                        String subProperty = "address"
                        def subtagKeySet = node[subProperty].keySet();
                        xmlMarkup.'Address'((subtagKeySet.getAt(0).toString()): node[subProperty][subtagKeySet.getAt(0)],
                                (subtagKeySet.getAt(1).toString()): node[subProperty][subtagKeySet.getAt(1)],
                                (subtagKeySet.getAt(2).toString()): node[subProperty][subtagKeySet.getAt(2)],
                                (subtagKeySet.getAt(3).toString()): node[subProperty][subtagKeySet.getAt(3)],
                                (subtagKeySet.getAt(4).toString()): node[subProperty][subtagKeySet.getAt(4)],
                        )
                    }

            }
        }
        //println(xmlWriter.toString())
        def file = new File(toXmlFile)
        file.createNewFile()
        file.write(xmlWriter.toString())
    }
}
