using System;
using CoreTweet;


namespace Excelion
{
    class TwiApiExec
    {
        private Tokens tokens;

        public TwiApiExec(Tokens t)
        {
            tokens = t;
        }

        public string fav(long FI)
        {
            try
            {
                Status favoritedStatus = tokens.Favorites.Create(id => FI);
                Console.WriteLine("fav exec");
                return ("お気に入りに追加しました -> " + FI);
            }
            catch (Exception ex)
            {
                Console.WriteLine("--------------fav errer-----------");
                Console.WriteLine(ex.Message);
                return (ex.Message);
            }
        }

        public string RT(long RI)
        {
            try
            {
                Status retweetedStatus = tokens.Statuses.Retweet(id => RI);
                Console.WriteLine("rt exec");
                return ("リツートしました -> " + RI);
            }
            catch (Exception ex)
            {
                Console.WriteLine("----------rt errer---------------");
                Console.WriteLine(ex);
                return ex.Message;
            }
        }

        public string Tweet(string tweet)
        {
            try
            {
                tokens.Statuses.Update(new { status = tweet });
                Console.WriteLine("tweet compreat!!");
                return ("ツイートしました -> " + tweet);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return (ex.Message);
            }
        }
    }
}
