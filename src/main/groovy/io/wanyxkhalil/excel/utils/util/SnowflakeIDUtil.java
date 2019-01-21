package io.wanyxkhalil.excel.utils.util;

/**
 * SnowflakeID生成器
 * 支持2^2=4个节点，每毫秒内同一个节点可以生成1<<20=1048576个Id
 */
class SnowflakeIDUtil {
    private static final SnowflakeIDUtil INSTANCE = new SnowflakeIDUtil();

    /**
     * 基准时间Thu Sep 01 00:00:00 CST 2016
     */
    private final static long START_STAMP = 1472659200000L;
    /**
     * 数据中心标志位数
     */
    private final static long DC_BIT = 1L;
    /**
     * 机器标识位数
     */
    private final static long MACHINE_BIT = 1L;
    /**
     * 序列号识位数
     */
    private final static long SEQUENCE_BIT = 20L;
    /**
     * 最大值
     */
    private final static long MAX_DC_NUM = ~(-1L << DC_BIT);
    /**
     * 机器ID最大值
     */
    private final static long MAX_MACHINE_NUM = ~(-1L << MACHINE_BIT);
    /**
     * 序列号ID最大值
     */
    private final static long MAX_SEQUENCE = ~(-1L << SEQUENCE_BIT);
    /**
     * 向左的位移
     */
    private final static long MACHINE_LEFT = SEQUENCE_BIT;
    private final static long DC_LEFT = SEQUENCE_BIT + MACHINE_LEFT;
    private final static long TIMESTAMP_LEFT = SEQUENCE_BIT + MACHINE_BIT + DC_BIT;

    private long dataCenterId;
    private long machineId;
    private long sequence = 0L;
    private long lastStamp = -1L;

    private SnowflakeIDUtil() {
        this.dataCenterId = 1L;
        this.machineId = 1L;
    }

    static SnowflakeIDUtil getInstance() {
        return INSTANCE;
    }

    /**
     * 产生Id
     */
    synchronized long nextId() {
        long currStamp = newStamp();
        if (currStamp < lastStamp) {
            throw new RuntimeException("Clock moved backwards.  Refusing to generate id");
        }

        // 相同毫秒内
        if (lastStamp == currStamp) {
            // 序列号自增
            sequence = (sequence + 1) & MAX_SEQUENCE;
            //同一毫秒的序列数已经达到最大，自旋等待到下一毫秒
            if (sequence == 0L) {
                currStamp = tailNextMillis();
            }
        } else {
            //不同毫秒内，序列号置为0
            sequence = 0L;
        }

        lastStamp = currStamp;

        // 时间戳部分
        return (currStamp - START_STAMP) << TIMESTAMP_LEFT
                // 数据中心部分
                | dataCenterId << DC_LEFT
                // 机器标识部分
                | machineId << MACHINE_LEFT
                // 序列号部分
                | sequence;
    }

    /**
     * 防止产生的时间比之前的时间还要小（由于NTP回拨等问题）,保持增量的趋势.
     */
    private long tailNextMillis() {
        long timestamp = newStamp();
        while (timestamp <= lastStamp) {
            timestamp = newStamp();
        }
        return timestamp;
    }

    /**
     * 获取当前的时间戳
     */
    private long newStamp() {
        return System.currentTimeMillis();
    }
}