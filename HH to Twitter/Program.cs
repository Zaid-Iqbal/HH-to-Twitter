using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tweetinvi;
using Tweetinvi.Models;
using Tweetinvi.Parameters;

namespace HH_to_Twitter
{
    class Program
    {
        public static IWorkbook Pwb = new XSSFWorkbook(@"C:\Users\email\Desktop\Hardware Hub\HHPosts.xlsx");
        public static ISheet Pws = Pwb.GetSheetAt(0);

        public static IWorkbook Twb = new XSSFWorkbook(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Twitter.xlsx");
        public static ISheet Tws = Twb.GetSheetAt(0);
        static void Main(string[] args)
        {
            String ID = getNextID();
            PublishTweet(getTweetBody(ID), ID);
            writeSheet(ID);
        }

        public static void PublishTweet(String msg, String ID)
        {
            Auth.SetUserCredentials(
               "qXnu7gz9p6CCFSkOmnZrvVwLj",
               "RewPba42pWihEuTNWXXazNhIHG1prDM6CO3kZX5ZiiTkV1NVAz",
               "1221234396080017409-pFxnlVAlJv0s30yJ5SdPnbvFaia2U7",
               "6EJqGZXnNdWvcUC1n890dA5yE8YSHtEuA25h1BE81R3PK");
            var user = User.GetAuthenticatedUser();

            if (msg.Length > 280)
            {
                MessageBox.Show("Tweet too long");
            }
            else
            {
                byte[] file1 = File.ReadAllBytes(@"C:\Users\email\Desktop\Hardware Hub\images\" + ID + ".png");
                var media = Upload.UploadBinary(file1);
                Tweet.PublishTweet(msg, new PublishTweetOptionalParameters
                {
                    Medias = new List<IMedia> { media }
                });
            }

        }

        public static string getTweetBody(String ID)
        {
            for (int x = 1; x <= Tws.LastRowNum; x++)
            {
                if (Tws.GetRow(x).Cells[1].ToString() == ID)
                {
                    String send = Tws.GetRow(x).Cells[5].ToString();
                    return send.Substring(send.IndexOf(":") + 2);
                }
            }
            MessageBox.Show("Excel.getNextID() Error: Body not found (Check if there are unsent tweets scheduled in Twitter.xlsx)");
            return "Not Found";
        }

        public static String getNextID()
        {
            for (int x = 0; x <= Tws.LastRowNum; x++)
            {
                IRow row = Tws.GetRow(x);
                if (row.Cells[2].ToString() == "No")
                {
                    String send = row.Cells[1].ToString();
                    Pwb.Close();
                    return send;
                }
            }
            MessageBox.Show("Excel.getNextID() Error: Unsent ID not found (Check if there are unsent tweets scheduled in Twitter.xlsx)");
            return "Not Found";

        }
        public static void writeSheet(String ID)
        {
            IWorkbook newWB = new XSSFWorkbook();
            ISheet newWS = newWB.CreateSheet();
            for (int x = 0; x <= Tws.LastRowNum; x++)
            {
                IRow row = Tws.GetRow(x);
                newWS.CreateRow(x);
                foreach (ICell cell in row)
                {
                    newWS.GetRow(x).CreateCell(cell.ColumnIndex).SetCellValue(cell.ToString());
                }
            }
            for (int x = 0; x <= newWS.LastRowNum; x++)
            {
                IRow row = newWS.GetRow(x);
                if (row.Cells[1].ToString() == ID)
                {
                    row.Cells[2].SetCellValue("Yes");
                    break;
                }
            }
            Twb.Close();
            newWB.Write(new FileStream(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Twitter.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite));
            newWB.Close();
            Twb = new XSSFWorkbook(@"C:\Users\email\Desktop\Hardware Hub\Twitter code files\Twitter.xlsx");
        }
    }
}
