namespace ExcelCalibrationAddin.Host.Recognition
{
    public static class FieldHeaderVocabulary
    {
        public static readonly string[] CalibrationItemTitleKeywords =
        {
            "示值误差", "基本误差", "引用误差", "重复性", "稳定性", "漂移", "响应时间",
            "报警功能", "报警动作值", "外观", "绝缘电阻", "零点漂移", "量程漂移"
        };

        public static readonly string[] StandardValueKeywords =
        {
            "标准值", "标准点", "标称值", "校准值", "参考值",
            "标准气浓度", "标气浓度", "标准浓度", "标准输入", "输入值", "标气浓度值",
            "标准", "标气"
        };

        public static readonly string[] SetpointValueKeywords =
        {
            "设定值", "设定", "给定值", "给定", "预设值", "预置值", "目标值"
        };

        public static readonly string[] ReferenceMeasurementValueKeywords =
        {
            "标准器测量值", "标准器示值", "标准器读数", "标准器实测值",
            "参考器测量值", "参考器示值", "参考器读数", "标准实际值",
            "实际标准值", "标准测量值"
        };

        public static readonly string[] MeasurementValueKeywords =
        {
            "测量值", "实测值", "读数", "读值", "示值", "仪器示值", "显示值", "检测值",
            "观测值", "采样值", "测得值", "测量", "测试值", "报警值", "动作值"
        };

        public static readonly string[] AverageValueKeywords =
        {
            "AVG", "Average", "平均", "平均值", "均值", "算术平均值"
        };

        public static readonly string[] ErrorValueKeywords =
        {
            "误差", "示值误差", "绝对误差", "相对误差", "引用误差", "基本误差",
            "重复性", "稳定性", "漂移", "重复性误差", "稳定性误差", "%FS"
        };

        public static readonly string[] TechnicalRequirementKeywords =
        {
            "技术要求", "允许误差", "允差", "最大允许误差", "MPE", "±MPE",
            "限值", "误差限", "指标要求", "报警动作值允许误差", "要求", "漂移"
        };

        public static readonly string[] RangeValueKeywords =
        {
            "量程", "范围", "测量范围", "满量程", "满程", "满刻度",
            "FS", "Full Scale", "Span"
        };

        public static readonly string[] UncertaintyKeywords =
        {
            "不确定度", "扩展不确定度", "U", "Urel", "urel", "k=2", "uncertainty"
        };

        public static readonly string[] ResultKeywords =
        {
            "结论", "判定", "P/F", "合格判定", "结果判定"
        };
    }
}
