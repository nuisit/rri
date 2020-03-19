using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PvdCalculator.Models;
using System.Text.RegularExpressions;

namespace PvdCalculator
{
    using Newtonsoft.Json;
    using PvdCalculator.Extensions;
    //using PvdCalculator.Models.Report;
    using System.Collections;
    using System.Windows.Forms.DataVisualization.Charting;
    using PvdCalculator.Models.Chart;

    public partial class Form1 : Form
    {
        private List<WorldBankLifespanModel> _worldBankLifespanList;

        private List<EmployeeModel> _employeeList;
        private List<PolicyRateModel> _policyRateList;

        private PolicyRateCriteriaModel _policyRateCriteria;
        private EmployeeCriteriaModel _employeeCriteria;
        private CriterionNumberModel _criterionNumber;

        /// <summary>
        /// อัตราเงินเฟ้อต่อปี (avg 1997-2017)
        /// </summary>
        public decimal InflationPercentage => 2.632M;
        public decimal Inflation => InflationPercentage / 100;
        /// <summary>
        /// ค่ากลาง (สีเหลือง) ของกราฟ
        /// </summary>
        public decimal AmberZonePercentage => 85;
        public decimal AmberZone => AmberZonePercentage / 100;

        public Form1()
        {
            InitializeComponent();
        }

        private void Recalculate()
        {
            if (_employeeList != null)
            {
                _policyRateList = GetPolicyRateList(_policyRateCriteria);
                _employeeList = AdjustImportData(_employeeList);

                GenerateChartRRI(_employeeList);
                GenerateChartZoneDistribution(_employeeList);
                GenerateChartMemberCanLive(_employeeList);
                GenerateChartAgeDistribution(_employeeList);
                GenerateChartAssetAllocationDistribution(_employeeList);
                GenerateChartContributionRateDistribution(_employeeList);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _worldBankLifespanList = JsonConvert.DeserializeObject<List<WorldBankLifespanModel>>(LoadJson("Data/worldbank-lifespan.json"));

            _criterionNumber = new CriterionNumberModel();
            _employeeCriteria = new EmployeeCriteriaModel();
            _policyRateCriteria = new PolicyRateCriteriaModel();

            _policyRateList = GetPolicyRateList(_policyRateCriteria);

            var costMonthlyAmountList = JsonConvert.DeserializeObject<List<ListItemModel>>(LoadJson("Data/costMonthlyAmount.json"));
            cbbCostMonthlyAmount.DataSource = costMonthlyAmountList;
            cbbCostMonthlyAmount.DisplayMember = "Text";
            cbbCostMonthlyAmount.ValueMember = "Value";
            cbbCostMonthlyAmount.SelectedIndex = 1;

            var lifespanList = JsonConvert.DeserializeObject<List<ListItemModel>>(LoadJson("Data/lifespan.json"));
            cbbLifespan.DataSource = lifespanList;
            cbbLifespan.DisplayMember = "Text";
            cbbLifespan.ValueMember = "Value";
            cbbLifespan.SelectedIndex = 4;

            var automaticEscalationList = JsonConvert.DeserializeObject<List<ListItemModel>>(LoadJson("Data/automaticEscalation.json"));
            cbbAutomaticEscalation.DataSource = automaticEscalationList;
            cbbAutomaticEscalation.DisplayMember = "Text";
            cbbAutomaticEscalation.ValueMember = "Value";

            var collectAdjustRateList = JsonConvert.DeserializeObject<List<ListItemModel>>(LoadJson("Data/collectAdjustRates.json"));
            cbbCollectAdjustRate.DataSource = collectAdjustRateList;
            cbbCollectAdjustRate.DisplayMember = "Text";
            cbbCollectAdjustRate.ValueMember = "Value";
            cbbCollectAdjustRate.SelectedIndex = 1;

            decimal.TryParse(((ListItemModel)cbbCostMonthlyAmount.SelectedItem).Value, out decimal costMonthlyAmount);
            _employeeCriteria.CostMonthlyAmount = costMonthlyAmount;
            int.TryParse(((ListItemModel)cbbLifespan.SelectedItem).Value, out int lifespan);
            _employeeCriteria.Lifespan = lifespan;
            _employeeCriteria.EmployersRatePercentage = ntxtEmployersRatePercentage.Value;
            _employeeCriteria.RetirementAge = decimal.ToInt32(ntxtRetirementAge.Value);
            _employeeCriteria.SalaryIncreaseRatePercentage = ntxtSalaryIncreaseRatePercentage.Value;
            _employeeCriteria.SalaryCap = ntxtSalaryCap.Value;

            _employeeCriteria.IsDoTargetDate = chkIsDoTargetDate.Checked;

            this.ntxtMarketPolicy.Value = _policyRateCriteria.MarketPolicyPercentage;
            this.ntxtBondPolicy.Value = _policyRateCriteria.BondPolicyPercentage;
            this.ntxtThaiStockPolicy.Value = _policyRateCriteria.ThaiStockPolicyPercentage;
            this.ntxtRealEstatePolicy.Value = _policyRateCriteria.RealEstatePolicyPercentage;
            this.ntxtForeignFundPolicy.Value = _policyRateCriteria.ForeignFundPolicyPercentage;
        }

        private void cbbCostMonthlyAmount_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal.TryParse(((ListItemModel)((ComboBox)sender).SelectedItem).Value, out decimal costMonthlyAmount);
            _employeeCriteria.CostMonthlyAmount = costMonthlyAmount;

            Recalculate();
        }

        private void cbbLifespan_SelectedIndexChanged(object sender, EventArgs e)
        {
            int.TryParse(((ListItemModel)((ComboBox)sender).SelectedItem).Value, out int lifespan);
            _employeeCriteria.Lifespan = lifespan;

            Recalculate();
        }

        private void ntxtEmployersRatePercentage_ValueChanged(object sender, EventArgs e)
        {
            _employeeCriteria.EmployersRatePercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void ntxtRetirementAge_ValueChanged(object sender, EventArgs e)
        {
            _employeeCriteria.RetirementAge = decimal.ToInt32(((NumericUpDown)sender).Value);

            Recalculate();
        }

        private void ntxtSalaryIncreaseRatePercentage_ValueChanged(object sender, EventArgs e)
        {
            _employeeCriteria.SalaryIncreaseRatePercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void ntxtSalaryCap_ValueChanged(object sender, EventArgs e)
        {
            _employeeCriteria.SalaryCap = ((NumericUpDown)sender).Value;

            Recalculate();
        }


        private void cbbAutomaticEscalation_SelectedIndexChanged(object sender, EventArgs e)
        {
            int.TryParse(((ListItemModel)((ComboBox)sender).SelectedItem).Value, out int automaticEscalation);
            _employeeCriteria.AutomaticEscalation = automaticEscalation;

            Recalculate();
        }

        private void cbbCollectAdjustRate_SelectedIndexChanged(object sender, EventArgs e)
        {
            decimal.TryParse(((ListItemModel)((ComboBox)sender).SelectedItem).Value, out decimal collectAdjustRate);
            _employeeCriteria.CollectAdjustRatePercentage = collectAdjustRate;

            Recalculate();
        }

        private void chkIsDoTargetDate_CheckedChanged(object sender, EventArgs e)
        {
            _employeeCriteria.IsDoTargetDate = ((CheckBox)sender).Checked;

            Recalculate();
        }

        private void ntxtMarketPolicy_ValueChanged(object sender, EventArgs e)
        {
            _policyRateCriteria.MarketPolicyPercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void ntxtBondPolicy_ValueChanged(object sender, EventArgs e)
        {
            _policyRateCriteria.BondPolicyPercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void ntxtThaiStockPolicy_ValueChanged(object sender, EventArgs e)
        {
            _policyRateCriteria.ThaiStockPolicyPercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void ntxtRealEstatePolicy_ValueChanged(object sender, EventArgs e)
        {
            _policyRateCriteria.RealEstatePolicyPercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void ntxtForeignFundPolicy_ValueChanged(object sender, EventArgs e)
        {
            _policyRateCriteria.ForeignFundPolicyPercentage = ((NumericUpDown)sender).Value;

            Recalculate();
        }

        private void GenerateChartRRI(List<EmployeeModel> list)
        {
            chartRRI.Series.Clear();

            if (list != null && 0 < list.Count)
            {
                var rriScore = (list.Sum((o) => o.Score) / list.Count) * 100;

                var series = new Series
                {
                    Name = "Series0",
                    ChartType = SeriesChartType.Bar,
                    IsVisibleInLegend = false
                };

                chartRRI.Series.Add(series);

                series.Points.AddXY("RRI Score", rriScore);
                series.Points[0].Label = rriScore.ToString("0.##") + "%";

                chartRRI.ChartAreas[0].AxisY.Minimum = 0;
                chartRRI.ChartAreas[0].AxisY.Maximum = 100;
                chartRRI.ChartAreas[0].AxisX.LineWidth = 0;
                chartRRI.ChartAreas[0].AxisY.LineWidth = 0;
                chartRRI.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
                chartRRI.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
                chartRRI.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chartRRI.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chartRRI.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                chartRRI.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                chartRRI.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
                chartRRI.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
                chartRRI.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                chartRRI.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;
                chartRRI.ChartAreas[0].BackColor = SystemColors.Control;
            }
        }

        private void GenerateChartZoneDistribution(List<EmployeeModel> list)
        {
            chartZoneDistribution.Series.Clear();

            if (list != null && 0 < list.Count)
            {
                int[] zoneLevels = new int[] { 0, 1, 2 };

                var series = new Series
                {
                    Name = "Series0",
                    ChartType = SeriesChartType.Pie,
                    IsVisibleInLegend = false
                };

                chartZoneDistribution.Series.Add(series);

                for (int i = 0; i < zoneLevels.Count(); i++)
                {
                    int zoneCount = list.Count((o) => o.ZoneLevel == zoneLevels[i]);
                    var item = new ZoneDistributionModel
                    {
                        ZoneLevel = zoneLevels[i],
                        Value = zoneCount,
                        ValuePercentage = ((decimal)zoneCount / list.Count) * 100
                    };

                    series.Points.AddXY(item.ZoneLevel, item.Value);
                    series.Points[i].Label = item.ValuePercentage.ToString("0.##") + "%";
                    series.Points[i].Color = item.Color;
                }
            }
        }

        private void GenerateChartMemberCanLive(List<EmployeeModel> list)
        {
            chartMemberCanLive.Series.Clear();

            if (list != null && 0 < list.Count)
            {
                var series = new Series
                {
                    Name = "Series0",
                    IsVisibleInLegend = false
                };

                chartMemberCanLive.Series.Add(series);

                int[] years = new int[] { 5, 10, 15, 20, 25, 30 };

                var deadYears = list.Max((o) => o.DeadYears);
                years = years.Where((o) => o <= deadYears - _employeeCriteria.RetirementAge).ToArray();
                for (int i = 0; i < years.Count(); i++)
                {
                    int count = list.Count((o) => years[i] < o.AgeOutOfMoney - _employeeCriteria.RetirementAge);
                    var m = new MemberCanLiveModel
                    {
                        Key = "<= " + years[i],
                        Value = count,
                        ValuePercentage = ((decimal)count / list.Count) * 100
                    };

                    series.Points.AddXY(m.Key, m.ValuePercentage);
                    series.Points[i].Label = m.ValuePercentage.ToString("0.##") + "%";

                }

                chartMemberCanLive.ChartAreas[0].AxisX.LineWidth = 0;
                chartMemberCanLive.ChartAreas[0].AxisY.LineWidth = 0;
                //chartMemberCanLive.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                chartMemberCanLive.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;
                chartMemberCanLive.ChartAreas[0].BackColor = SystemColors.Control;
            }
        }

        private void GenerateChartAgeDistribution(List<EmployeeModel> list)
        {
            chartAgeDistribution.Series.Clear();

            if (list != null && 0 < list.Count)
            {
                int[] zoneLevels = new int[] { 0, 1, 2 };
                List<RangeModel<int>> AgeRangeList = new List<RangeModel<int>>();
                AgeRangeList.Add(new RangeModel<int>
                {
                    Start = 20,
                    End = 29,
                });
                AgeRangeList.Add(new RangeModel<int>
                {
                    Start = 30,
                    End = 39,
                });
                AgeRangeList.Add(new RangeModel<int>
                {
                    Start = 40,
                    End = 49,
                });
                AgeRangeList.Add(new RangeModel<int>
                {
                    Start = 50,
                    End = 59,
                });

                for (int i = 0; i < zoneLevels.Count(); i++)
                {
                    var series = new Series
                    {
                        Name = "Series" + (i + 1),
                        IsVisibleInLegend = false
                    };
                    chartAgeDistribution.Series.Add(series);

                    for (int j = 0; j < AgeRangeList.Count(); j++)
                    {
                        var m = new AgeDistributionModel
                        {
                            Key = AgeRangeList[j].Start + " - " + AgeRangeList[j].End,
                            ZoneLevel = zoneLevels[i],
                            AgeRange = AgeRangeList[j],
                            Value = list?.Count((o) => o.ZoneLevel == zoneLevels[i] && AgeRangeList[j].Start <= o.Age && o.Age <= AgeRangeList[j].End) ?? 0
                        };

                        series.Points.AddXY(m.Key, m.Value);
                        series.Points[j].Label = m.Value.ToString();
                        series.Points[j].Color = m.Color;
                    }
                }

                chartAgeDistribution.ChartAreas[0].AxisX.LineWidth = 0;
                chartAgeDistribution.ChartAreas[0].AxisY.LineWidth = 0;
                //chartAgeDistribution.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                chartAgeDistribution.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;
                chartAgeDistribution.ChartAreas[0].BackColor = SystemColors.Control;
                chartAgeDistribution.ChartAreas[0].RecalculateAxesScale();
            }
        }

        private void GenerateChartAssetAllocationDistribution(List<EmployeeModel> list)
        {
            chartAssetAllocationDistribution.Series.Clear();

            if (list != null && 0 < list.Count)
            {
                int[] zoneLevels = new int[] { 0, 1, 2 };
                string[] assetAllocations = new string[] { "Target Date", "Fixed Income Oriented", "Mixed", "Equities Oriented" };

                List<AssetAllocationDistributionModel> assetAllocationDistributionList = new List<AssetAllocationDistributionModel>();
                for (int i = 0; i < zoneLevels.Count(); i++)
                {
                    var series = new Series
                    {
                        Name = "Series" + (i + 1),
                        IsVisibleInLegend = false
                    };

                    chartAssetAllocationDistribution.Series.Add(series);

                    for (int j = 0; j < assetAllocations.Count(); j++)
                    {
                        var m = new AssetAllocationDistributionModel
                        {
                            ZoneLevel = zoneLevels[i],
                            AssetAllocation = assetAllocations[j],
                            Key = assetAllocations[j],
                            Value = list?.Count((o) => o.ZoneLevel == zoneLevels[i] && o.Group == assetAllocations[j]) ?? 0
                        };

                        series.Points.AddXY(m.Key, m.Value);
                        series.Points[j].Label = m.Value.ToString();
                        series.Points[j].Color = m.Color;
                    }
                }

                chartAssetAllocationDistribution.ChartAreas[0].AxisX.LineWidth = 0;
                chartAssetAllocationDistribution.ChartAreas[0].AxisY.LineWidth = 0;
                //chartAssetAllocationDistribution.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;
                chartAssetAllocationDistribution.ChartAreas[0].BackColor = SystemColors.Control;
                chartAssetAllocationDistribution.ChartAreas[0].RecalculateAxesScale();
            }
        }

        private void GenerateChartContributionRateDistribution(List<EmployeeModel> list)
        {
            chartContributionRateDistribution.Series.Clear();

            if (list != null && 0 < list.Count)
            {
                int[] zoneLevels = new int[] { 0, 1, 2 };
                decimal[] contributionRates = new decimal[] { 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15 };

                for (int i = 0; i < zoneLevels.Count(); i++)
                {
                    var series = new Series
                    {
                        Name = "Series" + (i + 1),
                        IsVisibleInLegend = false
                    };

                    chartContributionRateDistribution.Series.Add(series);

                    for (int j = 0; j < contributionRates.Count(); j++)
                    {
                        var m = new ContributionRateDistributionModel
                        {
                            ZoneLevel = zoneLevels[i],
                            ContributionRate = contributionRates[j],
                            Key = contributionRates[j].ToString(),
                            Value = list?.Count((o) => o.ZoneLevel == zoneLevels[i] && o.CollectRatePercentage == contributionRates[j]) ?? 0
                        };

                        series.Points.AddXY(m.Key, m.Value);
                        series.Points[j].Label = m.Value.ToString();
                        series.Points[j].Color = m.Color;
                    }
                }

                chartContributionRateDistribution.ChartAreas[0].AxisX.LineWidth = 0;
                chartContributionRateDistribution.ChartAreas[0].AxisY.LineWidth = 0;
                //chartContributionRateDistribution.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisX.MinorGrid.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisY.MinorGrid.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisX.MajorTickMark.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisX.MinorTickMark.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].AxisY.MinorTickMark.Enabled = false;
                chartContributionRateDistribution.ChartAreas[0].BackColor = SystemColors.Control;
                chartContributionRateDistribution.ChartAreas[0].RecalculateAxesScale();
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            DataTable dt = default(DataTable);
            string filePath = default(string);

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Excel File Dialog";
                //fdlg.InitialDirectory = @"C:\";
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = openFileDialog.FileName;
                    dt = GetDataTableFromExcel(filePath);
                }
            }

            if (dt != null && 0 < dt.Rows.Count)
            {
                _employeeList = new List<EmployeeModel>();
                foreach (DataRow dr in dt.Rows)
                {
                    _employeeList.Add(new EmployeeModel
                    {
                        EmployeeCode = dr.Field<string>(0),
                        CurrentSalary = ParseCurrency(dr.Field<string>(1)),
                        Age = dr.Field<int>(2),
                        Gender = dr.Field<string>(3),
                        WorkYears = dr.Field<int?>(4),
                        RawCollectRate = ParsePercentage(dr.Field<string>(5)),
                        CollectRatePercentage = ParsePercentage(dr.Field<string>(5)),
                        EmployersRate = dr.Field<decimal?>(6),
                        PvdAmount = ParseCurrency(dr.Field<string>(7)),

                        //LifePartPolicyPercentage = ParsePercentage(dr.Field<string>(8)),
                        //MarketPolicyPercentage = ParsePercentage(dr.Field<string>(9)),
                        //BondPolicyPercentage = ParsePercentage(dr.Field<string>(10)),
                        //ThaiStockPolicyPercentage = ParsePercentage(dr.Field<string>(11)),
                        //RealEstatePolicyPercentage = ParsePercentage(dr.Field<string>(12)),
                        //ForeignFundPolicyPercentage = ParsePercentage(dr.Field<string>(13)),

                        RawLifePartPolicyPercentage = ParsePercentage(dr.Field<string>(8)),
                        RawMarketPolicyPercentage = ParsePercentage(dr.Field<string>(9)),
                        RawBondPolicyPercentage = ParsePercentage(dr.Field<string>(10)),
                        RawThaiStockPolicyPercentage = ParsePercentage(dr.Field<string>(11)),
                        RawRealEstatePolicyPercentage = ParsePercentage(dr.Field<string>(12)),
                        RawForeignFundPolicyPercentage = ParsePercentage(dr.Field<string>(13)),
                    });
                }

                Recalculate();
            }
        }

        private static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            if (!string.IsNullOrEmpty(path))
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(path))
                    {
                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();
                    DataTable tbl = new DataTable();
                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }
                    var startRow = hasHeader ? 2 : 1;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = tbl.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }

                    return tbl;
                }
            }

            return default(DataTable);
        }

        public static decimal ParseCurrency(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                return decimal.Parse(Regex.Match(value, @"-?\d{1,3}(,\d{3})*(\.\d+)?").Value);
            }

            return default(decimal);
        }

        public static decimal ParsePercentage(string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                return decimal.Parse(value.TrimEnd(new char[] { '%', ' ' }));
            }

            return default(decimal);
        }

        public List<PolicyRateModel> GetPolicyRateList(PolicyRateCriteriaModel criteria)
        {
            List<PolicyRateModel> policyRateList = new List<PolicyRateModel>();
            policyRateList.Add(new PolicyRateModel
            {
                MarketPolicyPercentage = criteria.MarketPolicyPercentage,
                BondPolicyPercentage = criteria.BondPolicyPercentage,
                ThaiStockPolicyPercentage = criteria.ThaiStockPolicyPercentage,
                RealEstatePolicyPercentage = criteria.RealEstatePolicyPercentage,
                ForeignFundPolicyPercentage = criteria.ForeignFundPolicyPercentage,
                AgeRange = new RangeModel<int> { Start = 20, End = 30 },
                MMPercentage = 0,
                EQPercentage = 50,
                REITsPercentage = 10,
                GlobalEQPercentage = 10
            });
            policyRateList.Add(new PolicyRateModel
            {
                MarketPolicyPercentage = criteria.MarketPolicyPercentage,
                BondPolicyPercentage = criteria.BondPolicyPercentage,
                ThaiStockPolicyPercentage = criteria.ThaiStockPolicyPercentage,
                RealEstatePolicyPercentage = criteria.RealEstatePolicyPercentage,
                ForeignFundPolicyPercentage = criteria.ForeignFundPolicyPercentage,
                AgeRange = new RangeModel<int> { Start = 31, End = 40 },
                MMPercentage = 0,
                EQPercentage = 40,
                REITsPercentage = 10,
                GlobalEQPercentage = 10
            });
            policyRateList.Add(new PolicyRateModel
            {
                MarketPolicyPercentage = criteria.MarketPolicyPercentage,
                BondPolicyPercentage = criteria.BondPolicyPercentage,
                ThaiStockPolicyPercentage = criteria.ThaiStockPolicyPercentage,
                RealEstatePolicyPercentage = criteria.RealEstatePolicyPercentage,
                ForeignFundPolicyPercentage = criteria.ForeignFundPolicyPercentage,
                AgeRange = new RangeModel<int> { Start = 41, End = 50 },
                MMPercentage = 0,
                EQPercentage = 40,
                REITsPercentage = 5,
                GlobalEQPercentage = 5
            });
            policyRateList.Add(new PolicyRateModel
            {
                MarketPolicyPercentage = criteria.MarketPolicyPercentage,
                BondPolicyPercentage = criteria.BondPolicyPercentage,
                ThaiStockPolicyPercentage = criteria.ThaiStockPolicyPercentage,
                RealEstatePolicyPercentage = criteria.RealEstatePolicyPercentage,
                ForeignFundPolicyPercentage = criteria.ForeignFundPolicyPercentage,
                AgeRange = new RangeModel<int> { Start = 51, End = 100 },
                MMPercentage = 10,
                EQPercentage = 15,
                REITsPercentage = 2.5M,
                GlobalEQPercentage = 2.5M
            });

            return policyRateList;
        }

        public string LoadJson(string path)
        {
            try
            {
                using (StreamReader r = new StreamReader(path, Encoding.UTF8))
                {
                    return r.ReadToEnd();
                }
            }
            catch
            {
                throw;
            }
        }

        private int GetDeadYears(EmployeeModel emp)
        {
            if (_employeeCriteria.Lifespan == -1)
            {
                var lifespan = _worldBankLifespanList?.FirstOrDefault((o) => o.Age == emp.Age && o.Gender == emp.Gender)?.Lifespan ?? 0;
                var deadYears = (int)Math.Ceiling(lifespan) + 5;
                return 65 < deadYears ? deadYears : 65;
            }
            else
            {
                return _employeeCriteria.Lifespan;
            }
        }

        public List<EmployeeModel> AdjustImportData(List<EmployeeModel> list)
        {
            return list?.Select((o) =>
            {
                var rate = o.RawCollectRate + _employeeCriteria.CollectAdjustRatePercentage;

                o.CollectRatePercentage = rate < 15 ? rate : 15;
                o.RetirementYear = _employeeCriteria.RetirementAge - o.Age;
                //o.DeadYears = _employeeCriteria.Lifespan;
                o.DeadYears = GetDeadYears(o);
                //o.RemainAge = _employeeCriteria.Lifespan - _employeeCriteria.RetirementAge;
                //o.RemainAge = o.DeadYears - _employeeCriteria.RetirementAge;
                o.RemainAge = 110 - _employeeCriteria.RetirementAge;
                o.CostAmontAt85 = _employeeCriteria.CostYearlyAmountWhenAfter85 * (decimal)Math.Pow(decimal.ToDouble(1 + Inflation), o.RetirementYear + _criterionNumber.BeforCriterion + 1);
                o.CostAmontAtRetirement = _employeeCriteria.CostYearlyAmountBetween60And85 * (decimal)Math.Pow(decimal.ToDouble(1 + Inflation), o.RetirementYear);

                if (_employeeCriteria.IsDoTargetDate)
                {
                    o.LifePartPolicyPercentage = 100;
                    o.MarketPolicyPercentage = 0;
                    o.BondPolicyPercentage = 0;
                    o.ThaiStockPolicyPercentage = 0;
                    o.RealEstatePolicyPercentage = 0;
                    o.ForeignFundPolicyPercentage = 0;
                }
                else
                {
                    o.LifePartPolicyPercentage = o.RawLifePartPolicyPercentage;
                    o.MarketPolicyPercentage = o.RawMarketPolicyPercentage;
                    o.BondPolicyPercentage = o.RawBondPolicyPercentage;
                    o.ThaiStockPolicyPercentage = o.RawThaiStockPolicyPercentage;
                    o.RealEstatePolicyPercentage = o.RawRealEstatePolicyPercentage;
                    o.ForeignFundPolicyPercentage = o.RawForeignFundPolicyPercentage;
                }

                o.PvdForecastList = GetPvdForecastList(o);
                o.RetirementCostForecastList = GetRetirementCostForecastList(o);

                var fundEndedAgeList = o.RetirementCostForecastList?.Where((x) => x.FundEndedAge != 0).ToList();
                o.OutOfMoneyYears = fundEndedAgeList != null && 0 < fundEndedAgeList.Count ? fundEndedAgeList.Min((x) => x.FundEndedAge) : 0;

                //o.ValueAtDead = o.RetirementCostForecastList?.FirstOrDefault((x) => x.Age == _employeeCriteria.Lifespan)?.FundedStatusActual ?? 0;
                //o.ValueAtDeadAmber = o.RetirementCostForecastList?.FirstOrDefault((x) => x.Age == _employeeCriteria.Lifespan)?.FundedStatusAmberZone ?? 0;

                o.ValueAtDead = o.RetirementCostForecastList?.FirstOrDefault((x) => x.Age == o.DeadYears)?.FundedStatusActual ?? 0;
                o.ValueAtDeadAmber = o.RetirementCostForecastList?.FirstOrDefault((x) => x.Age == o.DeadYears)?.FundedStatusAmberZone ?? 0;

                return o;
            }).ToList();
        }

        public List<PvdForecastModel> GetPvdForecastList(EmployeeModel employee)
        {
            List<PvdForecastModel> list = new List<PvdForecastModel>();
            decimal currentSalary = employee.CurrentSalary;
            decimal portfolioValueActual = 0;
            int automaticEscalation = 0;
            decimal collectRate = 0;
            decimal collectAdjustRate = 0;

            for (int i = 0; i <= _employeeCriteria.RetirementAge - employee.Age; i++)
            {
                if (0 < i)
                {
                    currentSalary = currentSalary + (currentSalary * _employeeCriteria.SalaryIncreaseRate);

                    if (_employeeCriteria.SalaryCap < currentSalary)
                    {
                        currentSalary = _employeeCriteria.SalaryCap;
                    }
                }

                decimal policySummary = (employee.MarketPolicy * _policyRateList?[0].MarketPolicy) + (employee.BondPolicy * _policyRateList?[0].BondPolicy) + (employee.ThaiStockPolicy * _policyRateList?[0].ThaiStockPolicy) + (employee.RealEstatePolicy * _policyRateList?[0].RealEstatePolicy) + (employee.ForeignFundPolicy * _policyRateList?[0].ForeignFundPolicy) ?? 0;

                decimal policyRateValue = 0;
                if (0 < employee.LifePartPolicy)
                {
                    policyRateValue = _policyRateList?.FirstOrDefault((x) => (x.AgeRange.Start <= employee.Age + i && employee.Age + i <= x.AgeRange.End)).Value ?? 0;
                }

                decimal returnActual = policySummary + policyRateValue;

                if (0 < _employeeCriteria.AutomaticEscalation)
                {
                    if (0 < i)
                    {
                        if (automaticEscalation < i)
                        {
                            collectAdjustRate = collectAdjustRate + 0.01M;
                            automaticEscalation = automaticEscalation + _employeeCriteria.AutomaticEscalation;

                            if (0.15M < collectAdjustRate)
                            {
                                collectAdjustRate = 0.15M;
                            }
                        }
                    }
                }

                if (0 < i)
                {
                    collectRate = employee.CollectRate + collectAdjustRate;
                    if (0.15M < collectRate)
                    {
                        collectRate = 0.15M;
                    }
                }
                else
                {
                    collectRate = employee.CollectRate;
                }

                decimal contribution = (_employeeCriteria.EmployersRate + collectRate) * currentSalary * 12;

                if (0 < i)
                {
                    portfolioValueActual = (portfolioValueActual * (1 + returnActual)) + contribution;
                }
                else
                {
                    portfolioValueActual = employee.PvdAmount + contribution;
                }

                list.Add(new PvdForecastModel
                {
                    Age = employee.Age + i,
                    WorkYear = i,
                    Salary = currentSalary,
                    ReturnActual = returnActual,
                    Contribution = contribution,
                    PortfolioValueActual = portfolioValueActual
                });
            }

            return list;
        }

        public List<RetirementCostForecastModel> GetRetirementCostForecastList(EmployeeModel employee)
        {
            decimal estimatedExp = 0;
            decimal fundedStatusActual = 0;
            decimal fundedStatusAmberZone = 0;

            List<RetirementCostForecastModel> list = new List<RetirementCostForecastModel>();

            for (int i = 1; i <= employee.RemainAge; i++)
            {
                if (i == 1)
                {
                    estimatedExp = employee.CostAmontAtRetirement * (1 + Inflation);
                }
                else if (i == _criterionNumber.BeforCriterion + 1)
                {
                    estimatedExp = employee.CostAmontAt85;
                }
                else
                {
                    estimatedExp = estimatedExp * (1 + Inflation);
                }

                if (i == 1)
                {
                    fundedStatusActual = employee.ValueAtRetirement - estimatedExp;
                    fundedStatusAmberZone = employee.ValueAtRetirement - (estimatedExp * AmberZone);
                }
                else
                {
                    fundedStatusActual = fundedStatusActual - estimatedExp;
                    fundedStatusAmberZone = fundedStatusAmberZone - (estimatedExp * AmberZone);
                }

                if (0 < fundedStatusActual)
                {
                    fundedStatusActual = fundedStatusActual + (fundedStatusActual * _policyRateCriteria.MarketPolicy);
                }

                if (0 < fundedStatusAmberZone)
                {
                    fundedStatusAmberZone = fundedStatusAmberZone + (fundedStatusAmberZone * _policyRateCriteria.MarketPolicy);
                }

                list.Add(new RetirementCostForecastModel
                {
                    Year = i,
                    Age = _employeeCriteria.RetirementAge + i,
                    EstimatedExp = estimatedExp,
                    FundedStatusActual = fundedStatusActual,
                    FundedStatusAmberZone = fundedStatusAmberZone,
                    FundEndedAge = 0 < fundedStatusActual ? 0 : i + employee.RetirementYear + employee.Age
                });
            }

            return list;
        }
    }
}
