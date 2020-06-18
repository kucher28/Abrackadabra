﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Net;

namespace ConsoleApp6
{
    class Program
    {
        static void Main(string[] args)
        {
            string url = "https://api.icndb.com/jokes/random";
            string response;

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();

            using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
            {
                response = streamReader.ReadToEnd();
            }
            ValueJSON valueJSON = JsonConvert.DeserializeObject<ValueJSON>(response);
            Console.WriteLine("Шутка: {0}", valueJSON.Value.Joke);

            Console.ReadKey();
        }
    }
}
