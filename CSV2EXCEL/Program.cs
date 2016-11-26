using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CSV2EXCEL {
    class Program {
        /// <summary>
        /// 引数は、最初に保存するExcelのパス、それ以降はCSVのパス（複数可）
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args) {
            //try {
                var outPath = args[0];
                var s = Settings.Load(Cs._settingsxml);
                var dic = new Dictionary<int, List<string>>();
                args.Skip(1).ToList().ForEach(csv => {
                    var tmp = csv.GetCSVData(s.Delimiter);
                    //設定されてる列だけ取得
                    tmp.ForEach(a => {
                        s.ExSettings.ForEach(es => {
                            if(!dic.ContainsKey(es.CSVColumn))
                                dic.Add(es.CSVColumn, new List<string>());
                            dic[es.CSVColumn].Add(a[es.CSVColumn]);
                        });
                    });
                });
                //保存パスが存在するときは削除
                if(File.Exists(outPath)) File.Delete(outPath);

                //Excel作成
                outPath.CreateExcel(dic, s);

                //ExitCodeは整数の場合成功（データの総数）、不の場合エラー
                Environment.ExitCode = dic.First().Value.Count;
            //} catch(Exception ex) {
            //    Cs.Log(ex.Message);
            //    Environment.ExitCode = -1;
            //}
        }
    }
}
