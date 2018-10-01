using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        private string originalDB = "";

        private bool listBoxSelected = false;

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
            
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods




        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            string practiceClasses = "";
            string firstPracticeClass = "";
            var selectedItems = listView1.SelectedItems;
            if (selectedItems.Count > 0)
            {
                foreach (ListViewItem selectedItem in selectedItems)
                {
                    practiceClasses = practiceClasses + selectedItem.SubItems[0].Text + "','";
                    if (string.IsNullOrEmpty(firstPracticeClass))
                        firstPracticeClass = selectedItem.SubItems[0].Text;
                }
                practiceClasses = practiceClasses.Remove(practiceClasses.Length -3, 3);

                

                if (string.IsNullOrEmpty(originalDB))
                    MessageBox.Show("Please select the original database", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    runSQLqueries(practiceClasses, firstPracticeClass, originalDB);
            }
            else
                MessageBox.Show("Please select at least one office code", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);






        }


        private void runSQLqueries(string pracClass, string firstcode, string origDB)
        {
            
            //run stored procedures required
            try
            {
                using (var conn = new SqlConnection("Data Source=" + textBoxSQLserver.Text + ";Initial Catalog=" + JurisDbName + ";User id=AthensDBO;Password=Athens29442385;"))
                {
                    string sql = "EXEC sp_msforeachtable @command1='ALTER TABLE ? NOCHECK CONSTRAINT all'";
                    using (var command = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        command.ExecuteNonQuery();
                    }
                }
                using (var conn = new SqlConnection("Data Source=" + textBoxSQLserver.Text + ";Initial Catalog=" + JurisDbName + ";User id=AthensDBO;Password=Athens29442385;"))
                {
                    string sql = "EXEC sp_MSforeachtable @command1='ALTER TABLE ? DISABLE TRIGGER ALL'";
                    using (var command = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex1)
            {
                MessageBox.Show("Error executing stored procedures. Details: " + ex1.Message);
            }

            string errorTable = "";
            int steps = 54;
            try
            {
                //run actual queries
                errorTable = "sysparam";
                string commandSQL = "Update " + JurisDbName + ".dbo.SysParam " +
                    "Set sptxtvalue=TxtValue " +
                    "from(select " + origDB + ".dbo.sysparam.sptxtvalue as TxtValue from  " + origDB + ".dbo.sysparam where spname='FldClient') SP " +
                    "where spname='FldClient'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 1, steps);

                commandSQL = "Update " + JurisDbName + ".dbo.SysParam " +
                     "Set sptxtvalue=TxtValue " +
                     "from(select " + origDB + ".dbo.sysparam.sptxtvalue as TxtValue from  " + origDB + ".dbo.sysparam where spname='FldMatter') SP " +
                     "where spname='FldMatter'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 2, steps);

                errorTable = "ProfitCenter";
                commandSQL = "Insert into " + JurisDbName + ".dbo.ProfitCenter(pcntrnbr, pcntrdesc) Select Pcntrnbr, Pcntrdesc from " + origDB + ".dbo.ProfitCenter";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 3, steps);

                errorTable = "PracticeClass";
                commandSQL = "Insert into " + JurisDbName + ".dbo.PracticeClass(PrctClsCode,PrctClsDesc) Select PrctClsCode,PrctClsDesc from " + origDB + ".dbo.PracticeClass " +
                    " where PrctClsCode in ('" + pracClass + "')";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 4, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "Select (select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.PracticeClass.prctclscode), 'Y' as SysCreated,3000 as DocClass,'R' as DocType,14,PrctclsDesc,PrctClsCode from " + JurisDbName + ".dbo.PracticeClass";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 5, steps);

                errorTable = "ActCode";
                commandSQL = "Insert into " + JurisDbName + ".dbo.ActivityCode(ActyCdCode,ActyCdDesc,ActyCdText) Select ActyCdCode,ActyCdDesc,ActyCdText from " + origDB + ".dbo.ActivityCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 6, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.ActivityCode.ActyCdCode),'Y' as SysCreated,6500 as DocClass,'R' as DocType,18,ActyCdDesc,ActyCdCode from " + JurisDbName + ".dbo.ActivityCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 7, steps);

                errorTable = "TaskCode";
                commandSQL = "Insert into " + JurisDbName + ".dbo.TaskCode(TaskCdCode,TaskCdDesc,TaskCdUseHrs,TaskCdUseRate,TaskCdUseAmt,TaskCdText) Select TaskCdCode,TaskCdDesc,TaskCdUseHrs,TaskCdUseRate,TaskCdUseAmt,TaskCdText from " + origDB + ".dbo.TaskCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 8, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.taskCode.TaskCdCode), 'Y' as SysCreated,2400 as DocClass,'R' as DocType,15,TaskCdDesc,TaskCdCode from " + JurisDbName + ".dbo.TaskCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 9, steps);

                if (checkBoxExpCodes.Checked)
                {
                    errorTable = "ExpCode";
                    commandSQL = "Insert into " + JurisDbName + ".dbo.ExpenseCode(ExpCdCode,ExpCdDesc ,ExpCdExpType,ExpCdTax1Exempt,ExpCdTax2Exempt,ExpCdTax3Exempt ,ExpCdText,ExpActive) Select ExpCdCode,ExpCdDesc ,ExpCdExpType,ExpCdTax1Exempt,ExpCdTax2Exempt,ExpCdTax3Exempt ,ExpCdText,ExpActive from " + origDB + ".dbo.ExpenseCode";
                    _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                    UpdateStatus("Updating Database", 10, steps);

                    commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                        "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.ExpenseCode.ExpCdCode), 'Y' as SysCreated,2900 as DocClass,'R' as DocType,19,ExpCdDesc,ExpCdCode from " + JurisDbName + ".dbo.ExpenseCode";
                    _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                    UpdateStatus("Updating Database", 11, steps);
                }

                errorTable = "BankAccounts";
                commandSQL = "Insert into " + JurisDbName + ".dbo.BankAccount(BnkCode,BnkDesc,BnkAcctType,BnkAcctNbr,BnkNextCheckNbr,BnkLastReconDate,BnkLastReconBal,BnkCheckLayout) " +
                    " Select BnkCode,BnkDesc,BnkAcctType,BnkAcctNbr,BnkNextCheckNbr,BnkLastReconDate,BnkLastReconBal,BnkCheckLayout from " + origDB + ".dbo.BankAccount ";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 12, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.BankAccount.BnkCode), 'Y' as SysCreated,6400 as DocClass,'R' as DocType,10,BnkDesc,BnkCode from " + JurisDbName + ".dbo.BankAccount";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 13, steps);

                errorTable = "BillFormat/Layout";
                commandSQL = "Insert into " + JurisDbName + ".dbo.BillFormat(BFCode,BFDesc,BFFontName,BFFontSize,BFTopMargin,BFLeftMargin,BFBottomMargin,BFRightMargin,BFHasCoverPage,BFHasExpAttachment,BFFirstPageSource,BFOtherPageSource,BFIncludeIfAR,BFIncludeIfNoDetails,BFAge1,BFAge2,BFAge3,BFAge4,BFStatus) Select BFCode,BFDesc,BFFontName,BFFontSize,BFTopMargin,BFLeftMargin,BFBottomMargin,BFRightMargin,BFHasCoverPage,BFHasExpAttachment,BFFirstPageSource,BFOtherPageSource,BFIncludeIfAR,BFIncludeIfNoDetails,BFAge1,BFAge2,BFAge3,BFAge4,BFStatus from " + origDB + ".dbo.BillFormat where bfcode not in (select bfcode from " + JurisDbName + ".dbo.BillFormat)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 14, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillFormatItem(BFICode,BFISection,BFISeq,BFIItemType,BFIItemValue,BFITop,BFILeft,BFIWidth,BFIHeight,BFIFontName,BFIFontSize,BFIBold,BFIItalic,BFIUnderline) Select BFICode,BFISection,BFISeq,BFIItemType,BFIItemValue,BFITop,BFILeft,BFIWidth,BFIHeight,BFIFontName,BFIFontSize,BFIBold,BFIItalic,BFIUnderline from " + origDB + ".dbo.BillFormatItem where bficode not in (select bfcode from " + JurisDbName + ".dbo.BillFormat)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 15, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillLayout(BLCode,BLDescription,BLLastModifiedBy,BLLastModifiedOn,BLLockedBy,BLSelectionOptions,BLPrintOptions,BLGUID,BLOptions) Select BLCode,BLDescription,BLLastModifiedBy,BLLastModifiedOn,BLLockedBy,BLSelectionOptions,BLPrintOptions,BLGUID,BLOptions from " + origDB + ".dbo.BillLayout where blcode not in (select blcode from " + JurisDbName + ".dbo.BillLayout)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 16, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillLayoutSection(BLSCode,BLSSegment,BLSSection,BLSName,BLSXML,BLSXSL,BLSOptions) Select BLSCode,BLSSegment,BLSSection,BLSName,BLSXML,BLSXSL,BLSOptions from " + origDB + ".dbo.BillLayoutSection where blscode not in (select blscode from " + JurisDbName + ".dbo.BillLayoutSection)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 17, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillLayoutSegment(BLSCode,BLSSegment,BLSName,BLSOptions) Select BLSCode,BLSSegment,BLSName,BLSOptions from " + origDB + ".dbo.BillLayoutSegment where blscode not in (select blscode from " + JurisDbName + ".dbo.BillLayoutSegment)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 18, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillLayoutTransformation(BLTCode,BLTSegment,BLTType,BLTXSL) Select BLTCode,BLTSegment,BLTType,BLTXSL from " + origDB + ".dbo.BillLayoutTransformation  where BLTCode not in (select BLTCode from " + JurisDbName + ".dbo.BillLayoutTransformation)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 19, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select  (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.BillFormat.BFCode),  'Y' as SysCreated,5500 as DocClass,'R' as DocType,46,BFDesc,BFCode from " + JurisDbName + ".dbo.BillFormat where bfcode<>'BF01'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 20, steps);

                errorTable = "OfficeCode";
                commandSQL = "Insert into " + JurisDbName + ".dbo.OfficeCode(OfcOfficeCode,OfcDesc,OfcAddrs1,OfcAddrs2,OfcAddrs3,OfcAddrs4,OfcAddrs5,OfcProfitCenter,OfcTaxActngMethod,OfcTax1OnFees,OfcTax1OnCshExp,OfcTax1OnNCshExp,OfcTax1OnSurChg,OfcTax1Description,OfcTax1Rate,OfcTax1MaxTax,OfcTax2OnFees,OfcTax2OnCshExp,OfcTax2OnNCshExp,OfcTax2OnSurChg,OfcTax2OnTax1,OfcTax2Description,OfcTax2Rate,OfcTax2MaxTax,OfcTax3OnFees,OfcTax3OnCshExp,OfcTax3OnNCshExp,OfcTax3OnSurChg,OfcTax3OnTax1,OfcTax3OnTax2,OfcTax3Description,OfcTax3Rate,OfcTax3MaxTax,OfcBankCode) Select OfcOfficeCode,OfcDesc,OfcAddrs1,OfcAddrs2,OfcAddrs3,OfcAddrs4,OfcAddrs5,OfcProfitCenter,OfcTaxActngMethod,OfcTax1OnFees,OfcTax1OnCshExp,OfcTax1OnNCshExp,OfcTax1OnSurChg,OfcTax1Description,OfcTax1Rate,OfcTax1MaxTax,OfcTax2OnFees,OfcTax2OnCshExp,OfcTax2OnNCshExp,OfcTax2OnSurChg,OfcTax2OnTax1,OfcTax2Description,OfcTax2Rate,OfcTax2MaxTax,OfcTax3OnFees,OfcTax3OnCshExp,OfcTax3OnNCshExp,OfcTax3OnSurChg,OfcTax3OnTax1,OfcTax3OnTax2,OfcTax3Description,OfcTax3Rate,OfcTax3MaxTax,OfcBankCode from " + origDB + ".dbo.OfficeCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 21, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.OfficeCode.OfcOfficeCode), 'Y' as SysCreated,2200 as DocClass,'R' as DocType,11,OfcDesc,OfcOfficeCode from " + JurisDbName + ".dbo.OfficeCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 22, steps);

                errorTable = "FeeSched";
                commandSQL = "Insert into " + JurisDbName + ".dbo.FeeSchedule(FeeSchCode,FeeSchDesc,FeeSchActive) Select FeeSchCode,FeeSchDesc,FeeSchActive from " + origDB + ".dbo.FeeSchedule where feeschcode in (select matfeesch from " + origDB + ".dbo.matter where matpracticeclass in ('" + pracClass + "')) or feeschcode in (select clifeesch from " + origDB + ".dbo.client where clipracticeclass in ('" + pracClass + "')) or " + origDB + ".dbo.FeeSChedule.feeschcode='STDR'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 23, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.FeeSchedule.FeeSchCode), 'Y' as SysCreated,2700 as DocClass,'R' as DocType,17,FeeSchDesc,FeeSchCode from " + JurisDbName + ".dbo.FeeSchedule";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 24, steps);

                errorTable = "ExpSched";
                commandSQL = "Insert into " + JurisDbName + ".dbo.ExpenseSchedule(ExpSchCode,ExpSchDesc) Select ExpSchCode,ExpSchDesc from " + origDB + ".dbo.ExpenseSchedule where  expschcode in (select matexpsch from " + origDB + ".dbo.matter where matpracticeclass in ('" + pracClass + "')) or expschcode in (select cliexpsch from " + origDB + ".dbo.client where clipracticeclass in ('" + pracClass + "')) or expschcode='STDR'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 25, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.ExpenseSchedule.ExpSchCode), 'Y' as SysCreated,4100 as DocClass,'R' as DocType,21,ExpSchDesc,ExpSchCode from " + JurisDbName + ".dbo.ExpenseSchedule";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 26, steps);

                errorTable = "PersType";
                commandSQL = "Insert into " + JurisDbName + ".dbo.PersonnelType(PrsTypCode,PrsTypDesc) Select PrsTypCode,PrsTypDesc from " + origDB + ".dbo.PersonnelType";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 27, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyT) " +
                    "select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.PersonnelType.PrsTypCode), 'Y' as SysCreated,2300 as DocClass,'R' as DocType,12,PrsTypDesc,prstypcode from " + JurisDbName + ".dbo.PersonnelType";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 28, steps);

                errorTable = "Exployee";
                commandSQL = "Insert into " + JurisDbName + ".dbo.Employee(EmpSysNbr,EmpID,EmpInitials,EmpPassword,EmpValidAsUser,EmpValidAsTkpr,EmpMnuPerm,EmpRptPerm,EmpOnLine,EmpPreferences,EmpPrsTyp,EmpSortSeq,EmpFeeTaxExempt,EmpBudHrsDay,EmpBudTargetRate,EmpNetID,EmpFirstName,EmpMiddleName,EmpLastName,EmpEmail) " +
                    " Select rank() over(order by " + origDB + ".dbo.Employee.EmpSysnbr) + 1,EmpID,EmpInitials,EmpPassword,EmpValidAsUser,EmpValidAsTkpr,EmpMnuPerm,EmpRptPerm,EmpOnLine,EmpPreferences,EmpPrsTyp,EmpSortSeq,EmpFeeTaxExempt,EmpBudHrsDay,EmpBudTargetRate,EmpNetID,EmpFirstName,EmpMiddleName,EmpLastName,EmpEmail from " + origDB + ".dbo.Employee" +
                    " where empsysnbr<>1 and (empsysnbr in (select tbdtkpr from " + origDB + ".dbo.timebatchdetail inner join  " + origDB + ".dbo.matter  on tbdmatter=matsysnbr where matpracticeclass in ('" + pracClass + "')) " +
                    " or empsysnbr in (Select clibillingatty from " + origDB + ".dbo.client where clipracticeclass in ('" + pracClass + "')) or empsysnbr in (select billtobillingatty from " + origDB + ".dbo.billto inner join " + origDB + ".dbo.matter on matbillto=billtosysnbr " +
                    " where matpracticeclass in ('" + pracClass + "')) or empsysnbr in (select morigatty from " + origDB + ".dbo.matorigatty inner join " + origDB + ".dbo.matter on morigmat=matsysnbr where matpracticeclass in ('" + pracClass + "'))" +
                    " or empsysnbr in (select corigatty from " + origDB + ".dbo.cliorigatty inner join " + origDB + ".dbo.client on corigcli=clisysnbr where clipracticeclass in ('" + pracClass + "'))" +
                    " or empsysnbr in (select MRTEmployeeID from " + origDB + ".dbo.MatterResponsibleTimekeeper inner join " + origDB + ".dbo.matter on MRTMatterID=matsysnbr where matpracticeclass in ('" + pracClass + "'))" +
                    " or empsysnbr in (select CRTEmployeeID from " + origDB + ".dbo.ClientResponsibleTimekeeper inner join " + origDB + ".dbo.client on CRTClientID=clisysnbr where clipracticeclass in ('" + pracClass + "')) )";
                int employees = _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 29, steps);

                if (employees == 0)
                {
                    MessageBox.Show("No Employees were added - check your SQL or selections. The application will now close", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    System.Environment.Exit(0);
                }

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyl) "
                    + " select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.Employee.EmpSysnbr), 'Y' as SysCreated,2600 as DocClass,'R' as DocType,13,EmpName,EmpSysNbr from " + JurisDbName + ".dbo.employee where empsysnbr<>1 and empvalidastkpr = 'Y'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyl) "
                    + " select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.Employee.EmpSysnbr), 'Y' as SysCreated,2500 as DocClass,'R' as DocType,52,EmpName,EmpSysNbr from " + JurisDbName + ".dbo.employee where empsysnbr<>1 and empvalidasuser = 'Y'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 30, steps);

                errorTable = "TkprRate";
                commandSQL = "Insert into " + JurisDbName + ".dbo.TkprRate(TKRFeeSch,TKREmp,TKRRate) " +
                    "Select FeeSchCode, EmpSysNbr, TkrRate " +
                    "from " + JurisDbName + ".dbo.Employee " +
                    "inner join (" +
                    "Select TKRFeeSch,EmpID as Tkpr,TKRRate from " + origDB + ".dbo.TkprRate " +
                    "inner join " + origDB + ".dbo.Employee on tkremp=empsysnbr) TRate on Tkpr=EmpID " +
                    "inner join " + JurisDbName + ".dbo.FeeSchedule on tkrfeesch=feeschcode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 31, steps);

                errorTable = "PersTypRate";
                commandSQL = "Insert into " + JurisDbName + ".dbo.PersTypRate(PTRFeeSch,PTRPrsTyp,PTRRate) Select PTRFeeSch,PTRPrsTyp,PTRRate from " + origDB + ".dbo.PersTypRate where ptrfeesch in (select feeschcode from  " + JurisDbName + ".dbo.FeeSchedule)";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 32, steps);

                errorTable = "Client";
                commandSQL = "Insert into " + JurisDbName + ".dbo.Client(CliSysnbr,CliCode,CliNickName,CliReportingName,CliSourceOfBusiness,CliPhoneNbr,CliFaxNbr,CliContactName,CliDateOpened,Cliofficecode,CliBillingAtty,CliPracticeClass,CliFeeSch,CliTaskCodeXref,CliExpSch,CliExpCodeXref,CliBillFormat,CliBillAgreeCode,CliFlatFeeIncExp,CliRetainerType,CliExpFreqCode,CliFeeFreqCode,CliBillMonth,CliBillCycle,CliExpThreshold,CliFeeThreshold,CliInterestPcnt,CliInterestDays,CliDiscountOption,CliDiscountPcnt,CliSurchargeOption,CliSurchargePcnt,CliTax1Exempt,CliTax2Exempt,CliTax3Exempt,CliBudgetOption,CliReqPhaseOnTrans,CliReqTaskCdOnTime,CliReqActyCdOnTime,CliReqTaskCdOnExp,CliPrimaryAddr,CliType,CliEditFormat,CliThresholdOption,CliRespAtty,CliBillingField01,CliBillingField02,CliBillingField03,CliBillingField04,CliBillingField05,CliBillingField06,CliBillingField07,CliBillingField08,CliBillingField09,CliBillingField10,CliBillingField11,CliBillingField12,CliBillingField13,CliBillingField14,CliBillingField15,CliBillingField16,CliBillingField17,CliBillingField18,CliBillingField19,CliBillingField20,CliCTerms,CliCStatus,CliCStatus2) " +
                    " Select rank() over (order by " + origDB + ".dbo.Client.CliSysnbr) as CliSys, CliCode,CliNickName,CliReportingName,CliSourceOfBusiness,CliPhoneNbr,CliFaxNbr,CliContactName,CliDateOpened,cliofficecode,Tkpr.EmpSysnbr,CliPracticeClass,CliFeeSch,CliTaskCodeXref,CliExpSch,CliExpCodeXref,CliBillFormat,CliBillAgreeCode,CliFlatFeeIncExp,CliRetainerType,CliExpFreqCode,CliFeeFreqCode,CliBillMonth,CliBillCycle,CliExpThreshold,CliFeeThreshold,CliInterestPcnt,CliInterestDays,CliDiscountOption,CliDiscountPcnt,CliSurchargeOption,CliSurchargePcnt,CliTax1Exempt,CliTax2Exempt,CliTax3Exempt,CliBudgetOption,CliReqPhaseOnTrans,CliReqTaskCdOnTime,CliReqActyCdOnTime,CliReqTaskCdOnExp,CliPrimaryAddr,CliType,CliEditFormat,CliThresholdOption,CliRespAtty,CliBillingField01,CliBillingField02,CliBillingField03,CliBillingField04,CliBillingField05,CliBillingField06,CliBillingField07,CliBillingField08,CliBillingField09,CliBillingField10,CliBillingField11,CliBillingField12,CliBillingField13,CliBillingField14,CliBillingField15,CliBillingField16,CliBillingField17,CliBillingField18,CliBillingField19,CliBillingField20,CliCTerms,CliCStatus,CliCStatus2 " +
                    " from " + origDB + ".dbo.Client" +
                    " inner join  " + origDB + ".dbo.Employee on  " + origDB + ".dbo.Employee.Empsysnbr=CliBillingATty" +
                    " inner join  " + JurisDbName + ".dbo.Employee Tkpr on Tkpr.EmpID= " + origDB + ".dbo.Employee.EmpID" +
                    " where " + origDB + ".dbo.Client.clicode in (select clicode from " + origDB + ".dbo.client where clipracticeclass in ('" + pracClass + "')) or " + origDB + ".dbo.Client.clisysnbr in (select matclinbr from " + origDB + ".dbo.matter where matpracticeclass in ('" + pracClass + "'))";
                int clients = _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 33, steps);

                if (clients == 0)
                {
                    MessageBox.Show("No clients were added - check your SQL or selections. The application will now close", "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    System.Environment.Exit(0);
                }

                commandSQL = "Insert into " + JurisDbName + ".dbo.DocumentTree(dtdocid,DTSystemCreated,DTDocClass,DTDocType,DTParentID,DTTitle,DTKeyl) " +
                    " select (Select  max(dtdocid) from " + JurisDbName + ".dbo.documenttree) + rank() over (order by " + JurisDbName + ".dbo.Client.Clisysnbr), 'Y' as SysCreated,4200 as DocClass,'R' as DocType,22,CliReportingName,CliSysnbr from " + JurisDbName + ".dbo.Client";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 34, steps);

                errorTable = "CliOrig";
                commandSQL = "Insert into " + JurisDbName + ".dbo.CliOrigAtty(COrigCli,COrigAtty,COrigPcnt) Select Cli.CliSysnbr,Emp.EmpSysNbr,COrigPcnt from " + origDB + ".dbo.CliOrigAtty" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.CliOrigAtty.COrigCli=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode" +
                    " Inner join " + origDB + ".dbo.Employee on " + origDB + ".dbo.CliOrigAtty.COrigAtty=" + origDB + ".dbo.Employee.EmpSysnbr" +
                    " Inner join " + JurisDbName + ".dbo.Employee Emp on  " + origDB + ".dbo.Employee.EmpID=Emp.EmpID";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 35, steps);

                errorTable = "CliResp";
                commandSQL = "Insert into " + JurisDbName + ".dbo.ClientResponsibleTimekeeper(CRTClientID,CRTEmployeeID,CRTPercent) " +
                    " Select Cli.CliSysnbr,Emp.EmpSysNbr,CRTPercent from " + origDB + ".dbo.ClientResponsibleTimekeeper" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.ClientResponsibleTimekeeper.CRTClientID=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode" +
                    " Inner join " + origDB + ".dbo.Employee on " + origDB + ".dbo.ClientResponsibleTimekeeper.CRTEmployeeID=" + origDB + ".dbo.Employee.EmpSysnbr" +
                    " Inner join " + JurisDbName + ".dbo.Employee Emp on  " + origDB + ".dbo.Employee.EmpID=Emp.EmpID";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 36, steps);

                commandSQL = "update " + JurisDbName + ".dbo.Client set CliPracticeClass = '" + firstcode + "' where CliPracticeClass not in ('" + pracClass + "')";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 37, steps);

                errorTable = "BillingAddy";
                commandSQL = "Insert into " + JurisDbName + ".dbo.BillingAddress(BilAdrSysnbr,BilAdrCliNbr,BilAdrUsageFlg,BilAdrNickName,BilAdrPhone,BilAdrFax,BilAdrContact,BilAdrName,BilAdrAddress,BilAdrCity,BilAdrState,BilAdrZip,BilAdrCountry,BilAdrType,BilAdrEmail) " +
                    " Select rank() over (order by " + origDB + ".dbo.BillingAddress.biladrsysnbr),Cli.CliSysNbr,BilAdrUsageFlg,BilAdrNickName,BilAdrPhone,BilAdrFax,BilAdrContact,BilAdrName,BilAdrAddress,BilAdrCity,BilAdrState,BilAdrZip,BilAdrCountry,BilAdrType,BilAdrEmail " +
                    " from " + origDB + ".dbo.BillingAddress" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.BillingAddress.BilAdrCliNbr=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 38, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillTo(BillToSysnbr,BillToCliNbr,BillToUsageFlg,BillToNickName,BillToBillingAtty,BillToBillFormat,BillToEditFormat,BillToRespAtty) " +
                    " Select rank() over (order by " + origDB + ".dbo.BillTo.billtosysnbr),Cli.CliSysnbr,BillToUsageFlg,BillToNickName,Emp.EmpSysnbr,BillToBillFormat,BillToEditFormat,BillToRespAtty from " + origDB + ".dbo.BillTo" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.BillTo.BillToCliNbr=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode" +
                    " Inner join " + origDB + ".dbo.Employee on " + origDB + ".dbo.BillTo.BillTobillingatty=" + origDB + ".dbo.Employee.EmpSysnbr" +
                    " Inner join " + JurisDbName + ".dbo.Employee Emp on " + origDB + ".dbo.Employee.EmpID=Emp.EmpID";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 39, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.BillCopy(BilCpyBillTo,BilCpyBilAdr,BilCpyComment,BilCpyNbrOfCopies,BilCpyPrintFormat,BilCpyEmailFormat,BilCpyExportFormat,BilCpyARFormat) " +
                    " Select BT.BillToSysNbr,BA.BilAdrSysnbr,BilCpyComment,BilCpyNbrOfCopies,BilCpyPrintFormat,BilCpyEmailFormat,BilCpyExportFormat,BilCpyARFormat " +
                    " from " + origDB + ".dbo.billcopy " +
                    " inner join " + origDB + ".dbo.billingaddress on " + origDB + ".dbo.billingaddress.biladrsysnbr=" + origDB + ".dbo.billcopy.bilcpybiladr" +
                    " inner join " + origDB + ".dbo.billto on " + origDB + ".dbo.billcopy.bilcpybillto=" + origDB + ".dbo.billto.billtosysnbr" +
                    " inner join " + origDB + ".dbo.client on " + origDB + ".dbo.client.clisysnbr=billtoclinbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on Cli.CliCode=" + origDB + ".dbo.client.CliCode" +
                    " inner join " + JurisDbName + ".dbo.BillTo BT on " + origDB + ".dbo.BillTo.BillToUsageFlg=BT.BillToUSageflg and " + origDB + ".dbo.BillTo.billtonickname=BT.BilltoNickName and BT.BillToCliNbr=Cli.CliSysNbr" +
                    " inner join " + JurisDbName + ".dbo.BillingAddress BA on " + origDB + ".dbo.BillingAddress.biladrusageflg=BA.biladrusageFlg " +
                    " and " + origDB + ".dbo.BillingAddress.BilAdrNickName=BA.BilAdrNickName and " +
                    origDB + ".dbo.BillingAddress.BilAdrName=BA.BilAdrName and Cli.CliSysnbr=BA.biladrclinbr " +
                    " where " + origDB + ".dbo.billcopy.bilcpybillto in (select matbillto from " + origDB + ".dbo.matter where matpracticeclass in ('" + pracClass + "'))";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 40, steps);

                errorTable = "Matter";
                commandSQL = "Insert into " + JurisDbName + ".dbo.Matter(MatSysNbr,MatCliNbr,MatBillTo,MatCode,MatNickName,MatReportingName,MatDescription,MatRemarks,MatPhoneNbr,MatFaxNbr,MatContactName,MatDateOpened,MatStatusFlag,MatLockFlag,MatDateClosed,MatOfficeCode,MatPracticeClass,MatFeeSch,MatTaskCodeXref,MatExpSch,MatExpCodeXref,MatQuickAction,MatBillAgreeCode,MatFlatFeeIncExp,MatRetainerType,MatFltFeeOrRetainer,MatExpFreqCode,MatFeeFreqCode,MatBillMonth,MatBillCycle,MatExpThreshold,MatFeeThreshold,MatInterestPcnt,MatInterestDays,MatDiscountOption,MatDiscountPcnt,MatSurchargeOption,MatSurchargePcnt,MatSplitMethod,MatSplitThreshold,MatSplitPriorAmtBld,MatBudgetOption,MatBudgetPhase,MatReqPhaseOnTrans,MatReqTaskCdOnTime,MatReqActyCdOnTime,MatReqTaskCdOnExp,MatTax1Exempt,MatTax2Exempt,MatTax3Exempt,MatDateLastWork,MatDateLastExp,MatDateLastBill,MatDateLastStmt,MatDateLastPaymt,MatLastPaymtAmt,MatARLastBill,MatPaySinceLastBill,MatAdjSinceLastBill,MatPPDBalance,MatVisionAddr,MatThresholdOption,MatType,MatBillingField01,MatBillingField02,MatBillingField03,MatBillingField04,MatBillingField05,MatBillingField06,MatBillingField07,MatBillingField08,MatBillingField09,MatBillingField10,MatBillingField11,MatBillingField12,MatBillingField13,MatBillingField14,MatBillingField15,MatBillingField16,MatBillingField17,MatBillingField18,MatBillingField19,MatBillingField20,MatCTerms,MatCStatus,MatCStatus2) " +
                    " Select rank() over (order by " + origDB + ".dbo.Matter.MatSysnbr) as MatSys,Cli.Clisysnbr,BT.BillToSysnbr,MatCode,MatNickName,MatReportingName,MatDescription,MatRemarks,MatPhoneNbr,MatFaxNbr,MatContactName,MatDateOpened,MatStatusFlag,MatLockFlag,MatDateClosed,MatOfficeCode,MatPracticeClass,MatFeeSch,MatTaskCodeXref,MatExpSch,MatExpCodeXref,MatQuickAction,MatBillAgreeCode,MatFlatFeeIncExp,MatRetainerType,MatFltFeeOrRetainer,MatExpFreqCode,MatFeeFreqCode,MatBillMonth,MatBillCycle,MatExpThreshold,MatFeeThreshold,MatInterestPcnt,MatInterestDays,MatDiscountOption,MatDiscountPcnt,MatSurchargeOption,MatSurchargePcnt,MatSplitMethod,MatSplitThreshold,MatSplitPriorAmtBld,MatBudgetOption,MatBudgetPhase,MatReqPhaseOnTrans,MatReqTaskCdOnTime,MatReqActyCdOnTime,MatReqTaskCdOnExp,MatTax1Exempt,MatTax2Exempt,MatTax3Exempt,MatDateLastWork,MatDateLastExp,MatDateLastBill,MatDateLastStmt,MatDateLastPaymt,MatLastPaymtAmt,MatARLastBill,MatPaySinceLastBill,MatAdjSinceLastBill,MatPPDBalance,MatVisionAddr,MatThresholdOption,MatType,MatBillingField01,MatBillingField02,MatBillingField03,MatBillingField04,MatBillingField05,MatBillingField06,MatBillingField07,MatBillingField08,MatBillingField09,MatBillingField10,MatBillingField11,MatBillingField12,MatBillingField13,MatBillingField14,MatBillingField15,MatBillingField16,MatBillingField17,MatBillingField18,MatBillingField19,MatBillingField20,MatCTerms,MatCStatus,MatCStatus2" +
                    " from " + origDB + ".dbo.Matter" +
                    " Inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.Matter.MatClinbr=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.Clicode=Cli.CliCode" +
                    " Inner join " + JurisDbName + ".dbo.BillTo BT on Cli.Clisysnbr=BT.BillToCliNbr" +
                    " Inner join " + origDB + ".dbo.BillTo on " + origDB + ".dbo.Client.CliSysnbr=" + origDB + ".dbo.BillTo.BillToCliNbr and Bt.billtousageflg= " + origDB + ".dbo.BillTo.billtousageflg  " +
                    " and " + origDB + ".dbo.BillTo.billtosysnbr=" + origDB + ".dbo.Matter.MatBillTo" +
                    "  and Bt.BillToNickName= " + origDB + ".dbo.BillTo.BillToNickName    and Bt.BillToBillFormat= " + origDB + ".dbo.BillTo.BillToBillFormat     and Bt.BillToEditFormat= " + origDB + ".dbo.BillTo.BillToEditFormat  " +
                    "  where " + origDB + ".dbo.Matter.matpracticeclass in ('" + pracClass + "')";
                int matter = _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 41, steps);

                errorTable = "MattOrig";
                commandSQL = "Insert into " + JurisDbName + ".dbo.MatOrigAtty(MOrigMat,MOrigAtty,MOrigPcnt) " +
                    " Select Mat.MatSysnbr,Emp.EmpSysnbr,MOrigPcnt from " + origDB + ".dbo.MatOrigAtty" +
                    " inner join " + origDB + ".dbo.Matter on " + origDB + ".dbo.MatOrigAtty.MorigMat=" + origDB + ".dbo.Matter.MatSysnbr" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.Matter.MatCliNbr=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode" +
                    " inner join " + JurisDbName + ".dbo.Matter Mat on Cli.CliSysnbr=Mat.MatClinbr and Mat.MatCode=" + origDB + ".dbo.Matter.MatCode" +
                    " Inner join " + origDB + ".dbo.Employee on " + origDB + ".dbo.MatOrigAtty.MOrigAtty=" + origDB + ".dbo.Employee.EmpSysnbr" +
                    " Inner join " + JurisDbName + ".dbo.Employee Emp on  " + origDB + ".dbo.Employee.EmpID=Emp.EmpID";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL); 
                UpdateStatus("Updating Database", 42, steps);

                errorTable = "MattResp";
                commandSQL = "Insert into " + JurisDbName + ".dbo.MatterResponsibleTimekeeper(MRTMatterID,MRTEmployeeID,MRTPercent) Select MRTMatterID,MRTEmployeeID,MRTPercent from " + origDB + ".dbo.MatterResponsibleTimekeeper" +
                    " inner join " + origDB + ".dbo.Matter on " + origDB + ".dbo.MatterResponsibleTimekeeper.MRTMatterid=" + origDB + ".dbo.Matter.MatSysnbr" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.Matter.MatCliNbr=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode" +
                    " inner join " + JurisDbName + ".dbo.Matter Mat on Cli.CliSysnbr=Mat.MatClinbr and Mat.MatCode=" + origDB + ".dbo.Matter.MatCode" +
                    " Inner join " + origDB + ".dbo.Employee on " + origDB + ".dbo.MatterResponsibleTimekeeper.MRTEmployeeID=" + origDB + ".dbo.Employee.EmpSysnbr" +
                    " Inner join " + JurisDbName + ".dbo.Employee Emp on  " + origDB + ".dbo.Employee.EmpID=Emp.EmpID";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 43, steps);

                errorTable = "Notes";
                commandSQL = "Insert into " + JurisDbName + ".dbo.ClientNote(CNClient,CNNoteIndex,CNObject,CNNoteText,CNNoteObject) Select Cli.CliSysnbr,CNNoteIndex,CNObject,CNNoteText,CNNoteObject from " + origDB + ".dbo.ClientNote" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.ClientNote.CNClient=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 44, steps);

                commandSQL = "Insert into " + JurisDbName + ".dbo.MatterNote(MNMatter,MNNoteIndex,MNObject,MNNoteText,MNNoteObject) Select Mat.MatSysnbr,MNNoteIndex,MNObject,MNNoteText,MNNoteObject from " + origDB + ".dbo.MatterNote" +
                    " inner join " + origDB + ".dbo.Matter on " + origDB + ".dbo.MatterNote.MNmatter=" + origDB + ".dbo.Matter.MatSysnbr" +
                    " inner join " + origDB + ".dbo.Client on " + origDB + ".dbo.Matter.MatCliNbr=" + origDB + ".dbo.Client.CliSysnbr" +
                    " inner join " + JurisDbName + ".dbo.Client Cli on " + origDB + ".dbo.Client.CliCode=Cli.CliCode" +
                    " inner join " + JurisDbName + ".dbo.Matter Mat on Cli.CliSysnbr=Mat.MatClinbr and Mat.MatCode=" + origDB + ".dbo.Matter.MatCode";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 45, steps);

                errorTable = "Sysparam2";
                commandSQL = "update " + JurisDbName + ".dbo.sysparam " +
                    "set spnbrvalue=NbrValue from (select max(Clisysnbr) as NBrValue from " + JurisDbName + ".dbo.client) SysP where spname='LastSysNbrClient'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 46, steps);

                if (matter > 0)
                {
                    commandSQL = "update " + JurisDbName + ".dbo.sysparam " +
                        " set spnbrvalue=NbrValue from (select max(matsysnbr) as NBrValue from " + JurisDbName + ".dbo.matter) SysP where spname='LastSysNbrMatter'";
                    _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                }
                UpdateStatus("Updating Database", 47, steps);

                commandSQL = "update " + JurisDbName + ".dbo.sysparam " +
                    " set spnbrvalue=NbrValue from (select max(empsysnbr) as NBrValue from " + JurisDbName + ".dbo.employee) SysP where spname='LastSysNbrEmp'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 48, steps);

                commandSQL = "update " + JurisDbName + ".dbo.sysparam " +
                    " set spnbrvalue=NbrValue from (select max(biladrsysnbr) as NBrValue from " + JurisDbName + ".dbo.billingaddress) SysP where spname='LastSysNbrBillAddress'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 49, steps);

                commandSQL = "update " + JurisDbName + ".dbo.sysparam " +
                    " set spnbrvalue=NbrValue from (select max(billTosysnbr) as NBrValue from " + JurisDbName + ".dbo.billto) SysP where spname='LastSysNbrBillTo'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 50, steps);

                commandSQL = "update Juris9999003.dbo.sysparam set SpTxtValue = '12,N,N,N,N,N,1,N,N,Y,0,Y,N,N' where spname='CfgMiscOpts'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", 51, steps);

                commandSQL = "update " + JurisDbName + ".dbo.sysparam " +
                    " set spnbrvalue=NbrValue from (select max(DTDocID) as NBrValue from " + JurisDbName + ".dbo.documenttree) SysP where spname='LastSysNbrDocTree'";
                _jurisUtility.ExecuteNonQueryCommand(0, commandSQL);
                UpdateStatus("Updating Database", steps, steps);
            }
            catch (Exception ex2)
            {
                MessageBox.Show(errorTable + " : " + ex2.Message);
            }

            try
            {
                using (var conn = new SqlConnection("Data Source=" + textBoxSQLserver.Text + ";Initial Catalog=" + JurisDbName + ";User id=AthensDBO;Password=Athens29442385;"))
                {
                    string sql = "EXEC sp_msforeachtable @command1='ALTER TABLE ? CHECK CONSTRAINT ALL'";
                    using (var command = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        command.ExecuteNonQuery();
                    }
                }
                using (var conn = new SqlConnection("Data Source=" + textBoxSQLserver.Text + ";Initial Catalog=" + JurisDbName + ";User id=AthensDBO;Password=Athens29442385;"))
                {
                    string sql = "EXEC sp_MSforeachtable @command1='ENABLE TRIGGER ALL ON ?'";
                    using (var command = new SqlCommand(sql, conn))
                    {
                        conn.Open();
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex1)
            {
                MessageBox.Show("Error executing stored procedures. Details: " + ex1.Message);
            }



        }



        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            var selectedItems = listView1.SelectedItems;
            if (selectedItems.Count > 0)
                DoDaFix();
            else
                MessageBox.Show("Please ensure you have selected a SQL Server, an original database and at least one Practice Class", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }


        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxSelected)
            {
                originalDB = listBox1.GetItemText(listBox1.SelectedItem);
                bool isJurisDB = false;
                try
                {
                    var connString = "Data Source=" + textBoxSQLserver.Text + ";Initial Catalog=" + originalDB + ";User id=AthensDBO;Password=Athens29442385;";

                    SqlConnection conn = new SqlConnection(connString);
                    SqlDataAdapter da = new SqlDataAdapter();
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "select OfcOfficeCode, OfcDesc from officecode";
                    da.SelectCommand = cmd;
                    DataSet ds = new DataSet();

                    conn.Open();
                    da.Fill(ds);
                    conn.Close();
                    isJurisDB = true;
                }
                catch (Exception ex3)
                {
                    MessageBox.Show("That does not appear to be a Juris database", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    isJurisDB = false;
                    originalDB = "";
                }

                if (isJurisDB)
                {
                    originalDB = listBox1.GetItemText(listBox1.SelectedItem);

                    var connString = "Data Source=" + textBoxSQLserver.Text + ";Initial Catalog=" + originalDB + ";User id=AthensDBO;Password=Athens29442385;";

                    SqlConnection conn = new SqlConnection(connString);
                    SqlDataAdapter da = new SqlDataAdapter();
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "select PrctClsCode, PrctClsDesc from PracticeClass";
                    da.SelectCommand = cmd;
                    DataSet ds = new DataSet();

                    conn.Open();
                    da.Fill(ds);
                    conn.Close();

                    listView1.View = View.Details;
                    listView1.Columns.Add("Practice Class", 85);
                    listView1.Columns.Add("Practice Class Desc", 300);

                    foreach (DataRow row in ds.Tables[0].Rows)
                    {
                        listView1.Items.Add(new ListViewItem(new string[] { row[0].ToString(), row[1].ToString() }));
                    }
                }
            }
        }

        private void listBox1_Click(object sender, EventArgs e)
        {
            listBoxSelected = true;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            var connString = "Data Source=" + textBoxSQLserver.Text + ";Integrated Security=SSPI;";

            using (SqlConnection c = new SqlConnection(connString))
            {
                c.Open();

                // use a SqlAdapter to execute the query
                using (SqlDataAdapter a = new SqlDataAdapter("SELECT name from sys.databases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb')", c))
                {
                    // fill a data table
                    var t = new DataTable();
                    a.Fill(t);

                    // Bind the table to the list box
                    listBox1.DisplayMember = "name";
                    listBox1.ValueMember = "name";
                    listBox1.DataSource = t;
                }
            }
        }


    }
}
