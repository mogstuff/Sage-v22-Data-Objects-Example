using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SageSDO
{
    class Program
    {
        private const string sAccDataPath = @"C:\ProgramData\Sage\Accounts\2016\Company.000\ACCDATA";
        // NB To Free up Accounts data folder
        // C:\ProgramData\Sage\Accounts\2016\Company.000\ACCDATA\ delete file QUEUE.DTA
        

        static void Main(string[] args)
        {

              int iSalesPaymentSplit = 0;
      

            // get split on Invoice 5 and read through the info(5,6,7)
            SageDataObject220.SDOEngine oSDO = new SageDataObject220.SDOEngine();
            SageDataObject220.WorkSpace oWS;
            SageDataObject220.SalesRecord oSalesRecord;
            SageDataObject220.HeaderData oHeaderData;
            SageDataObject220.SplitData oSplitData;
            //  SageDataObject220.InvoiceRecord oInvoiceRecord;
            SageDataObject220.SplitData oPaymentSplitData;
            SageDataObject220.SplitData oCreditNoteSplitData;
            SageDataObject220.SplitData oSalesPaymentSplitData;


            string szDataPath;
            oWS = (SageDataObject220.WorkSpace)oSDO.Workspaces.Add("Example");

            szDataPath = oSDO.SelectCompany("C:\\ProgramData\\Sage\\Accounts\\2016");


            try
            {
                oWS.Connect(szDataPath, "manager", "", "MORGAN TECH");

                oSalesRecord = (SageDataObject220.SalesRecord)oWS.CreateObject("SalesRecord");
                string sAccountRef = "A001";
                oSalesRecord.Fields.Item("ACCOUNT_REF").Value = sAccountRef;


                // Get the Customer Record
                if (oSalesRecord.Find(false))
                {

                    // Move to top of Headers List for this Customer
                    oHeaderData = (SageDataObject220.HeaderData)oSalesRecord.Link;
                    oHeaderData.MoveFirst();

                    // Get No Of Splits
                    int iSplitCount = oHeaderData.Fields.Item("NO_OF_SPLIT").Value;
                    
               
                    while (oHeaderData.MoveNext())
                    {
                        // type of transaction
                        //1. Sales Invoice, 
                        //2. Sales Credit, 
                        //3. Sales Receipt, 
                        //4 Sales Receipt on Account
                        //24 Sales Payment
                        sbyte bType = oHeaderData.Fields.Item("TYPE").Value;
                        
                        // Get Credit Note based on INV_REF 4
                        #region Credit Note

                   
                        if (oHeaderData.Fields.Item("INV_REF").Value == "4")
                        {

                            // REF 4, Split 10 5.39, 11 2.40
                            oCreditNoteSplitData = oHeaderData.Link;
                            oCreditNoteSplitData.MoveFirst();

                            string itemDetails = oCreditNoteSplitData.Fields.Item("DETAILS").Value;
                            double netAmount = oCreditNoteSplitData.Fields.Item("NET_AMOUNT").Value;
                            double taxAmount = oCreditNoteSplitData.Fields.Item("TAX_AMOUNT").Value;
                            int transNo = oCreditNoteSplitData.Fields.Item("TRAN_NUMBER").Value;
                            int uniqueRef = oCreditNoteSplitData.Fields.Item("UNIQUE_REF").Value;

                            string message = string.Format("CRN Details : {0}  Net : {1} VAT {2} TransNo {3} Unique Ref {4}", itemDetails, netAmount, taxAmount, transNo, uniqueRef);
                            Console.WriteLine(message);
                            Console.ReadLine();

                            while (oCreditNoteSplitData.MoveNext())
                            {
                                itemDetails = oCreditNoteSplitData.Fields.Item("DETAILS").Value;
                                netAmount = oCreditNoteSplitData.Fields.Item("NET_AMOUNT").Value;
                                taxAmount = oCreditNoteSplitData.Fields.Item("TAX_AMOUNT").Value;
                                transNo = oCreditNoteSplitData.Fields.Item("TRAN_NUMBER").Value;
                                uniqueRef = oCreditNoteSplitData.Fields.Item("UNIQUE_REF").Value;

                                message = string.Format("CRN Details : {0}  Net : {1} VAT {2} TransNo {3} Unique Ref {4}", itemDetails, netAmount, taxAmount, transNo, uniqueRef);
                                
                                Console.WriteLine(message);
                                Console.ReadLine();

                            }

                        }

                        #endregion


                        // get Sales Payment (Refund)
                        #region Sales Payment (Refund)

                        //  if (oHeaderData.Fields.Item("INV_REF").Value == "12" || bType == 24)
                        if (bType == 24)
                        {

                            oSalesPaymentSplitData = oHeaderData.Link;
                            int intSalesPaymentNo = oHeaderData.Fields.Item("FIRST_SPLIT").Value;

                            iSalesPaymentSplit = intSalesPaymentNo;

                            string paymentItemDetails = oSalesPaymentSplitData.Fields.Item("DETAILS").Value;
                            double netAmountPaid = oSalesPaymentSplitData.Fields.Item("NET_AMOUNT").Value;
                            double taxAmountPaid = oSalesPaymentSplitData.Fields.Item("TAX_AMOUNT").Value;
                            int transNoPayment = oSalesPaymentSplitData.Fields.Item("TRAN_NUMBER").Value;
                            int uniqueRefPayment = oSalesPaymentSplitData.Fields.Item("UNIQUE_REF").Value;
                            string paymentMessage = string.Format("sales payment - REFUND found record no {0} net {1} tax {2}  trans no {3} unique ref {4}  ", intSalesPaymentNo, netAmountPaid, taxAmountPaid,
                                transNoPayment, uniqueRefPayment);

                            Console.WriteLine(paymentMessage);
                            Console.ReadLine();

                        }

                        #endregion


                        // allocate the sales payment against the credit note
                        #region Allocate

                        // allocate SP 12 against CR 10 (split 10 5.39 , 11 2.40)
                        // iSalesPaymentSplit  - first split of Sales Payment
                        if (oHeaderData.Fields.Item("INV_REF").Value == "4")
                        {


                            SageDataObject220.TransactionPost oPost = (SageDataObject220.TransactionPost)oWS.CreateObject("TransactionPost");

                            // REF 4, Split 10 5.39, 11 2.40
                            oCreditNoteSplitData = oHeaderData.Link;
                            oCreditNoteSplitData.MoveFirst();

                            string itemDetails = oCreditNoteSplitData.Fields.Item("DETAILS").Value;
                            double netAmount = oCreditNoteSplitData.Fields.Item("NET_AMOUNT").Value;
                            double taxAmount = oCreditNoteSplitData.Fields.Item("TAX_AMOUNT").Value;
                            int transNo = oCreditNoteSplitData.Fields.Item("TRAN_NUMBER").Value;
                            int uniqueRef = oCreditNoteSplitData.Fields.Item("UNIQUE_REF").Value;


                            // allocate first crnote line
                            //iSalesPaymentSplit
                            iSalesPaymentSplit = 12;

                            double amount = netAmount + taxAmount;
                            if (oPost.AllocatePayment(transNo, iSalesPaymentSplit, amount, System.DateTime.Now)) 
                            { 
                                Console.WriteLine("first amount allocated");
                            }


                            string message = string.Format("CRN Details : {0}  Net : {1} VAT {2} TransNo {3} Unique Ref {4}", itemDetails, netAmount, taxAmount, transNo, uniqueRef);
                            Console.WriteLine(message);
                            Console.ReadLine();

                            while (oCreditNoteSplitData.MoveNext())
                            {
                                itemDetails = oCreditNoteSplitData.Fields.Item("DETAILS").Value;
                                netAmount = oCreditNoteSplitData.Fields.Item("NET_AMOUNT").Value;
                                taxAmount = oCreditNoteSplitData.Fields.Item("TAX_AMOUNT").Value;
                                transNo = oCreditNoteSplitData.Fields.Item("TRAN_NUMBER").Value;
                                uniqueRef = oCreditNoteSplitData.Fields.Item("UNIQUE_REF").Value;

                                 amount = netAmount + taxAmount;
                                if (oPost.AllocatePayment(transNo, iSalesPaymentSplit, amount, System.DateTime.Now))
                                {
                                    Console.WriteLine("allocated");
                                }




                                message = string.Format("CRN Details : {0}  Net : {1} VAT {2} TransNo {3} Unique Ref {4}", itemDetails, netAmount, taxAmount, transNo, uniqueRef);

                                Console.WriteLine(message);
                                
                                
                                Console.ReadLine();

                            }

                        }



                        #endregion

                        // Get Payment Header based on INV_REF
                        #region Payment

                        if (oHeaderData.Fields.Item("INV_REF").Value == "PAYINVOICE2")
                        {
                            oPaymentSplitData = oHeaderData.Link;
                            // Get the FIRST_SPLIT for the payment, this is needed for allocations                    
                            int intPaymentSplitNo = oHeaderData.Fields.Item("FIRST_SPLIT").Value;

                            string paymentItemDetails = oPaymentSplitData.Fields.Item("DETAILS").Value;
                            double netAmountPaid = oPaymentSplitData.Fields.Item("NET_AMOUNT").Value;
                            double taxAmountPaid = oPaymentSplitData.Fields.Item("TAX_AMOUNT").Value;
                            int transNoPayment = oPaymentSplitData.Fields.Item("TRAN_NUMBER").Value;
                            int uniqueRefPayment = oPaymentSplitData.Fields.Item("UNIQUE_REF").Value;
                            string paymentMessage = string.Format("payment found record no {0} net {1} tax {2}  trans no {3} unique ref {4}  ", intPaymentSplitNo, netAmountPaid, taxAmountPaid,
                                transNoPayment, uniqueRefPayment);

                            Console.WriteLine(paymentMessage);
                            Console.ReadLine();

                        }

                        #endregion

                        // Get Invoice based on INV_REF           
                        #region Invoice

                        if (oHeaderData.Fields.Item("INV_REF").Value == "2")
                        {
                            oSplitData = oHeaderData.Link;
                            oSplitData.MoveFirst();

                            string itemDetails = oSplitData.Fields.Item("DETAILS").Value;
                            double netAmount = oSplitData.Fields.Item("NET_AMOUNT").Value;
                            double taxAmount = oSplitData.Fields.Item("TAX_AMOUNT").Value;
                            int transNo = oSplitData.Fields.Item("TRAN_NUMBER").Value;
                            int uniqueRef = oSplitData.Fields.Item("UNIQUE_REF").Value;

                            string message = string.Format("Details : {0}  Net : {1} VAT {2} TransNo {3} Unique Ref {4}", itemDetails, netAmount, taxAmount, transNo, uniqueRef);
                            Console.WriteLine(message);
                            Console.ReadLine();

                            while (oSplitData.MoveNext())
                            {
                                itemDetails = oSplitData.Fields.Item("DETAILS").Value;
                                netAmount = oSplitData.Fields.Item("NET_AMOUNT").Value;
                                taxAmount = oSplitData.Fields.Item("TAX_AMOUNT").Value;
                                transNo = oSplitData.Fields.Item("TRAN_NUMBER").Value;
                                uniqueRef = oSplitData.Fields.Item("UNIQUE_REF").Value;

                                message = string.Format("Details : {0}  Net : {1} VAT {2} TransNo {3} Unique Ref {4}", itemDetails, netAmount, taxAmount, transNo, uniqueRef);
                                Console.WriteLine(message);
                                Console.ReadLine();

                            }

                        }

                        #endregion


                    };


                }


                Console.ReadLine();

                oWS.Disconnect();

            }
            catch (Exception ex)
            {

                Console.WriteLine("SDO Generated the Following Error: \n\n" + ex.Message, "Error!");

            }



        }
    }
}
