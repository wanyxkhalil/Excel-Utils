package io.wanyxkhalil.excel.utils

import io.wanyxkhalil.excel.utils.testdomain.DomainGroovy
import io.wanyxkhalil.excel.utils.testdomain.DomainJava
import spock.lang.Specification

class ExcelWriterSpec extends Specification {

    def "生成Excel--java"() {
        setup:
        def list = new ArrayList<DomainJava>()

        1.upto(10000) {
            def v = new DomainJava()
            v.field1 = "第 $it 列"
            v.field2 = it.intValue()
            v.field3 = it.doubleValue()
            v.field4 = new BigDecimal(it.toString())

            list += v
        }

        when:
        def path = ExcelWriter.builder()
                .fileName()
                .sheet(list)
                .build()

        then:
        println path
    }


    def "生成Excel--groovy"() {
        setup:
        def list = new ArrayList<DomainGroovy>()

        1.upto(10000) {
            def v = new DomainGroovy()
            v.field1 = "第 $it 列"
            v.field2 = it.intValue()
            v.field3 = it.doubleValue()
            v.field4 = new BigDecimal(it.toString())

            list += v
        }

        when:
        def path = ExcelWriter.builder()
                .fileName()
                .sheet(list)
                .build()

        then:
        println path
    }

}
