using System;
using System.Linq;

using CoreTweet;
using CoreTweet.Streaming;

using System.Threading;


using Excel = Microsoft.Office.Interop.Excel;

namespace Excelion
{
    class Program
    {
        static void Main(string[] args)
        {
            string AK, AS, AT, ATS;

            AK = keys.AK;
            AS = keys.AS;

            //カレントディレクトリ
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
            Console.WriteLine(stCurrentDir);

            try
            {
                string excelName = stCurrentDir + "\\Excelion.xlsm";

                Excel.Application xlApplication = new Excel.Application();

                //対象エクセルを開く
                Excel.Workbook WorkBook = xlApplication.Workbooks.Open(excelName);
                Console.WriteLine("start excelion stream");

                //シート１をアクティブ
                Excel.Worksheet sheet1 = WorkBook.Sheets[1];
                Console.WriteLine("Excelion open");

                Excel.Worksheet sheet2 = WorkBook.Sheets[2];

                //フラグ開放
                sheet2.get_Range("H10").Value = 0;

                //使うとこ指定
                Excel.Range rangeTweet = sheet1.get_Range("C5", "D22");
                Excel.Range idousaki = sheet1.get_Range("C7", "D24");

                Excel.Range tweetId = sheet2.get_Range("B6", "B14");
                Excel.Range tweetId_idousaki = sheet2.get_Range("B7", "B15");

                Excel.Range SetId = sheet1.get_Range("C5");
                Excel.Range SetName = sheet1.get_Range("C6");
                Excel.Range SetTweet = sheet1.get_Range("D5");
                Excel.Range SetStId = sheet2.get_Range("B6");

                Excel.Range favId = sheet2.get_Range("G6");
                Excel.Range rtId = sheet2.get_Range("H6");
                Excel.Range Tweet = sheet2.get_Range("E9");

                Excel.Range favFlag = sheet2.get_Range("G7");
                Excel.Range rtFlag = sheet2.get_Range("H7");
                Excel.Range TwFlag = sheet2.get_Range("E10");

                Excel.Range News = sheet1.get_Range("D3");

                //初起動
                if (sheet2.get_Range("E13").Value != 1)
                {
                    getToken(sheet2, AK, AS);
                }
                //エクセル前面表示
                xlApplication.Visible = true;

                //トークン
                AT = sheet2.get_Range("C2").Value;
                ATS = sheet2.get_Range("C3").Value;

                //シート１表示
                sheet1.Select();

                Console.WriteLine("stream起動中");

                //こっからCoreTweet
                var t = Tokens.Create
                    (
                    AK
                    , AS
                    , AT
                    , ATS
                    );

                //t.Statuses.Update(status => "hello");

                //API操作インスタンス生成
                TwiApiExec TwitterAPI = new TwiApiExec(t);

                //publish stream
                var streamRx = t.Streaming.UserAsObservable();
                var stream = t.Streaming.StartObservableStream(StreamingType.User).Publish();

                //action
                Action<StatusMessage> printStatusHome = (message) =>
                {
                    try
                    {
                        var status = (message as StatusMessage).Status;
                        idousaki.Value = rangeTweet.Value;
                        tweetId_idousaki.Value = tweetId.Value;
                        SetId.Value = status.User.ScreenName;
                        SetName.Value = status.User.Name;
                        SetTweet.Value = (String)status.Text;
                        SetStId.Value = status.Id.ToString(); //文字列として打ち込まないと丸め誤差る
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("nettowa-kuera-");
                        Console.WriteLine(ex.Message);
                    }
                };

                //subscribe actions for event
                streamRx.OfType<StatusMessage>()
                    .Subscribe(
                    printStatusHome,
                    onError: exception => News.Value = (exception.ToString())
                    );

                //connect stream
                var connection = streamRx.Connect();


                //エクセルが閉じられるまで
                while (xlApplication.Visible)
                {
                    //End
                    if (sheet2.get_Range("H10").Value == 1)
                    {
                        //保存
                        WorkBook.Save();
                        //workbookを閉じる
                        WorkBook.Close();
                        //エクセルを閉じる
                        xlApplication.Quit();
                        //抜ける
                        break;
                    }

                    //fav
                    if (favFlag.Value == 1)
                    {
                        News.Value = TwitterAPI.fav(long.Parse(favId.Value2));
                        favId.Value = "";
                        favFlag.Value = 0;
                    }

                    //rt
                    if (rtFlag.Value == 1)
                    {
                        News.Value = TwitterAPI.RT(long.Parse(rtId.Value2));
                        rtId.Value = "";
                        rtFlag.Value = 0;
                    }

                    //tweet
                    if (TwFlag.Value == 1)
                    {
                        News.Value = TwitterAPI.Tweet(Tweet.Value2);
                        TwFlag.Value = 0;
                        Tweet.Value = "";
                    }


                    //適切な数字が分からんから勘で置いてる
                    Thread.Sleep(1000);
                }
                //close connection
                Console.WriteLine("close");
                connection.Dispose();
                Thread.Sleep(30);
            }
            catch (Exception ex)
            {
                Console.WriteLine("なんか分からんエラー");
                Console.WriteLine(ex.Message);
                Thread.Sleep(5000);
            }

        }

        static void getToken(Excel.Worksheet sheet, string AS, string AK)
        {
            try
            {

                var s = OAuth.Authorize(AS, AK);

                System.Diagnostics.Process.Start(s.AuthorizeUri.AbsoluteUri);

                Console.Write("PINコードを入力してください : ");
                string PIN = Console.ReadLine();
                Tokens tokens = s.GetTokens(PIN);

                sheet.get_Range("C2").Value = tokens.AccessToken;
                sheet.get_Range("C3").Value = tokens.AccessTokenSecret;
                sheet.get_Range("D1").Value = tokens.ScreenName;
                sheet.get_Range("E13").Value = 1;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
