using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using JTone.Com.ClosedXml.Excel;

namespace JTone.Com.ClosedXml.Test
{
    [TestClass]
    public class ExcelHelperTest
    { 
        [TestMethod]
        public void BuildHeaderTest()
        {
            var line = new List<DesignTaskGroupReportOutDto> { new DesignTaskGroupReportOutDto
            {
                ProjectName = "CD-2013388",
                CheckResultStageResults = new List<CheckStageResult>
                {
                    new CheckStageResult
                    {
                        QuestionCheckStage = QuestionCheckStageEnum.SelfCheck,
                        CheckEachStageResults = new List<CheckEachStageResult>
                        {
                            new CheckEachStageResult
                            {
                                CheckGradeName = "轻微一般",
                                ResultDesc = "-",
                            },
                            new CheckEachStageResult
                            {
                                CheckGradeName = "重要重大",
                                ResultDesc = "-",
                            },
                        }
                    }
                }
            }};

            var excelHelper = new ExcelHelper<DesignTaskGroupReportOutDto>("测试表", line);
            excelHelper.Export();
        }
    }
}
