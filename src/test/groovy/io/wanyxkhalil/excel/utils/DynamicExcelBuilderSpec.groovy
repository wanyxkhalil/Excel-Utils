package io.wanyxkhalil.excel.utils

import io.wanyxkhalil.excel.utils.dynamic.GroovyDynamicData
import spock.lang.Specification

class DynamicExcelBuilderSpec extends Specification {

    def "生成动态数据的Excel"() {
        setup:
        def list = new ArrayList<GroovyDynamicData>(10000)

        1.upto(10000) {

            def v = new GroovyDynamicData()

            v.propertySet("字符串", "第 $it 列")
            v.propertySet("Integer", it.intValue())
            v.propertySet("Double", it.doubleValue())
            v.propertySet("BigDecimal", new BigDecimal(it.toString()))

            list += v
        }

        when:
        def path = DynamicExcelUtil.builder()
                .fileName("fd")
                .sheet("1", list)
                .sheet("2", list[0,500])
                .build()
                .write2File()

        then:
        println path
    }
}
