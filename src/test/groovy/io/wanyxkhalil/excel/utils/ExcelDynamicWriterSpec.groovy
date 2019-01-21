package io.wanyxkhalil.excel.utils

import io.wanyxkhalil.excel.utils.dynamicwriter.DynamicData
import spock.lang.Specification

class ExcelDynamicWriterSpec extends Specification {

    def "生成动态数据的Excel"() {
        setup:
        def list = new ArrayList<DynamicData>(10000)

        1.upto(10000) {

            def v = new DynamicData()

            v.propertySet("字符串", "第 $it 列")
            v.propertySet("Integer", it.intValue())
            v.propertySet("Double", it.doubleValue())
            v.propertySet("BigDecimal", new BigDecimal(it.toString()))

            list += v
        }

        when:
        def path = ExcelDynamicWriter.builder()
                .fileName("fd")
                .sheet("1", list)
                .sheet("2", list[0, 500])
                .build()

        then:
        println path
    }
}
