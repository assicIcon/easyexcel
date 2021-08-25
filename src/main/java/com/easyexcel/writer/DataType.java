package com.easyexcel.writer;

import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.util.Date;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author wrh
 * @date 2021/5/13
 */
public enum DataType {

    /**
     * String.class
     */
    STRING(String.class) {
        @Override
        public Object transferType(Object object) {
            return object.toString();
        }
    },

    /**
     * java.util.Date.class
     */
    DATE(Date.class) {
        @Override
        public Object transferType(Object object) {
            if (object instanceof Date) {
                return object;
            }
            if (object instanceof Double) {
                return DateUtil.getJavaDate((Double) object);
            }
            if (object instanceof String) {
//                return DateTimeUtil.strToDate((String) object);
                // todo DateTimeFormat annotation
            }
            return object;
        }
    },

    /**
     * Integer.class
     */
    INT(Integer.class) {
        @Override
        public Object transferType(Object object) {
            if (object instanceof Number) {
                return ((Number) object).intValue();
            }
            if (object instanceof String) {
                return Integer.valueOf((String) object);
            }
            return object;
        }
    },

    /**
     * Long.class
     */
    LONG(Long.class) {
        @Override
        public Object transferType(Object object) {
            if (object instanceof Number) {
                return ((Number) object).longValue();
            }
            if (object instanceof String) {
                return Long.valueOf((String) object);
            }
            return object;
        }
    },

    /**
     * Float.class
     */
    FLOAT(Float.class) {
        @Override
        public Object transferType(Object object) {
            if (object instanceof Number) {
                return ((Number) object).floatValue();
            }
            if (object instanceof String) {
                return Float.valueOf((String) object);
            }
            return object;
        }
    },

    /**
     * Double.class
     */
    DOUBLE(Double.class) {
        @Override
        public Object transferType(Object object) {
            if (object instanceof Number) {
                return ((Number) object).doubleValue();
            }
            if (object instanceof String) {
                return Double.valueOf((String) object);
            }
            return object;
        }
    },

    /**
     * BigDecimal.class
     */
    BIG_DECIMAL(BigDecimal.class) {
        @Override
        public Object transferType(Object object) {
            if (object instanceof Number) {
                return new BigDecimal(object.toString());
            }
            if (object instanceof String) {
                return new BigDecimal(object.toString());
            }
            return object;
        }
    };
    private final Type type;

    public Type getType() {
        return type;
    }

    DataType(Type type) {
        this.type = type;
    }

    public Object transferType(Object object) {
        throw new RuntimeException("not support");
    }

    private static final Map<Type, DataType> DATA_TYPE_MAP;

    static {
        DATA_TYPE_MAP = Stream.of(values()).collect(Collectors.toMap(DataType::getType, e -> e));
    }

    public static DataType of(Type type) {
        return DATA_TYPE_MAP.get(type);
    }

}
