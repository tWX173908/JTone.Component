using System.Collections.Generic;
using System.ComponentModel;


namespace JTone.Com.ClosedXml.Test
{
    /// <summary>
    /// 设计任务组报表
    /// </summary>
    internal class DesignTaskGroupReportOutDto
    {
        /// <summary>
        /// 项目名称
        /// </summary>
        [Description("项目名称")]
        public string ProjectName { get; set; }

        /// <summary>
        /// 成果检查验证信息
        /// </summary>
        [Description("图审问题及验证率")]
        public List<CheckStageResult> CheckStageResults { get; set; }

        /// <summary>
        /// 任务-进行中
        /// </summary>
        [Description("进行中")]
        public string TaskDoing { get; set; }

        /// <summary>
        /// 任务-已完成
        /// </summary>
        [Description("已提交")]
        public string TaskFinished { get; set; }

        /// <summary>
        /// 任务-确认完成
        /// </summary>
        [Description("已完成")]
        public string TaskConfirmFinish { get; set; }
    }


    /// <summary>
    /// 成果检查验证信息
    /// </summary>
    internal class CheckStageResult
    {
        /// <summary>
        /// 成果检查阶段
        /// </summary>
        [DisplayName]
        public QuestionCheckStageEnum QuestionCheckStage { get; set; }

        /// <summary>
        /// 阶段信息
        /// </summary>
        public List<CheckEachStageResult> CheckEachStageResults { get; set; } = new List<CheckEachStageResult>();
    }


    /// <summary>
    /// 各阶段成果检查验证信息
    /// </summary>
    internal class CheckEachStageResult
    {
        /// <summary>
        /// 问题等级
        /// </summary>
        [DisplayName]
        public string CheckGradeName { get; set; }

        /// <summary>
        /// 问题状态
        /// </summary>
        public string ResultDesc { get; set; }
    }


    /// <summary>
    /// 问题检查阶段枚举
    /// </summary>
    internal enum QuestionCheckStageEnum
    {
        [Description("自校")]
        SelfCheck = 1,

        [Description("验证")]
        Check = 2,

        [Description("抽查")]
        SpotCheck = 4
    }


    /// <summary>
    /// 设计任务角色枚举
    /// </summary>
    internal enum DesignTaskRoleEnum
    {
        [Description("设计师")]
        Execute = 0,

        [Description("协助人")]
        AssistExecute = 1,

        [Description("专负")]
        Confirm = 2,

        [Description("协助确认人")]
        AssistConfirm = 3,

        [Description("设计师")]
        Members = 4,

        [Description("设总")]
        ExecuteManager = 5,

        [Description("方案")]
        Scheme = 6,

        [Description("校对")]
        Proofreading = 7,

        [Description("校核")]
        Check = 8,

        [Description("审核")]
        Vetted = 9,

        [Description("审定")]
        Validation = 10,

        [Description("汇总")]
        Total = 999,
    }
}
