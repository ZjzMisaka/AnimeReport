using ClosedXML.Excel;
using GlobalObjects;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace AnalyzeCode
{
    public class Analyze
    {
        public class Anime 
        {
            public AnimeType animeType;
            public Season season;
            public string year;
            public string name;
            public string origName;
            public string productionCompany;
            
            public Status status;
            public bool planToWatch;
            public int score;
            public string comment;
        }
        
        public enum Status { NeverWatched, Watching, Watched, GaveUp };
        
        public enum Season { Winter, Spring, Summer, Autumn, None };
        public static Dictionary<Season, List<string>> SeasonDic = new Dictionary<Season, List<string>>() 
        {
            { Season.Winter, new List<string>(){ "1月－3月"} },
            { Season.Spring, new List<string>(){ "4月－6月"} },
            { Season.Summer, new List<string>(){ "7月－9月"} },
            { Season.Autumn, new List<string>(){ "10月－12月"} }
        };
    
        public enum TitleType { Date, Name, OrigName, ProductionCompany };
        public static Dictionary<TitleType, List<string>> TitleDic = new Dictionary<TitleType, List<string>>() 
        {
            { TitleType.Date, new List<string>(){ "日"} },
            { TitleType.Name, new List<string>(){ "作品名", "中文译名"} },
            { TitleType.OrigName, new List<string>(){ "原名"} },
            { TitleType.ProductionCompany, new List<string>(){ "动画制作", "制作公司"} }
        };
        
        public enum AnimeType { TV, OVA, Movie, WEB, None };
        public static Dictionary<AnimeType, List<string>> AnimeTypeDic = new Dictionary<AnimeType, List<string>>() 
        {
            { AnimeType.TV, new List<string>(){ "电视动画"} },
            { AnimeType.OVA, new List<string>(){ "OVA"} },
            { AnimeType.Movie, new List<string>(){ "电影"} },
            { AnimeType.WEB, new List<string>(){ "WEB", "网络"} }
        };
        
        public bool IsBlank(IXLWorksheet sheet, int row, int col)
        {
            return sheet.Cell(row, col).CachedValue.ToString() == "";
        }
        
        /// <summary>
        /// すべての分析の前に呼び出されます
        /// </summary>
        /// <param name="param">着信パラメータ</param>
        /// <param name="globalObject">グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。</param>
        /// <param name="allFilePathList">分析されるすべてのファイルパスのリスト</param>
        /// <param name="isExecuteInSequence">順番実行するかどうか</param>
        public void RunBeforeAnalyzeSheet(Param param, ref Object globalObject, List<string> allFilePathList, bool isExecuteInSequence)
        {
            Output.IsSaveDefaultWorkBook = false;
        }

        /// <summary>
        /// シートを分析する
        /// </summary>
        /// <param name="param">着信パラメータ</param>
        /// <param name="sheet">分析するシート</param>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="globalObject">グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。</param>
        /// <param name="isExecuteInSequence">順番実行するかどうか</param>
        /// <param name="invokeCount">この分析関数が呼び出された回数</param>
        public void AnalyzeSheet(Param param, IXLWorksheet sheet, string filePath, ref Object globalObject, bool isExecuteInSequence, int invokeCount)
        {
            List<Anime> animeList = new List<Anime>();
        
            string year = sheet.Name.Substring(0, 4);
            
            int nowRow = 1;
            AnimeType nowAnimeType = AnimeType.None;
            Season nowSeason = Season.None;
            
            int nameCol = -1;
            int origNameCol = -1;
            int productionCompanyCol = -1;
            
            while (true)
            {
                if (IsBlank(sheet, nowRow, 3) && IsBlank(sheet, nowRow, 4) && !sheet.Cell(nowRow, 3).IsMerged() &&  !sheet.Cell(nowRow, 4).IsMerged())
                {
                    break;
                }
                
                if (!IsBlank(sheet, nowRow, 1))
                {
                    nameCol = -1;
                    origNameCol = -1;
                    productionCompanyCol = -1;
                    
                    string typeStr = sheet.Cell(nowRow, 1).CachedValue.ToString();
                    nowAnimeType = AnimeType.None;
                    foreach (string str in AnimeTypeDic[AnimeType.TV])
                    {
                        if (typeStr.Contains(str))
                        {
                            nowAnimeType = AnimeType.TV;
                            break;
                        }
                    }
                    foreach (string str in AnimeTypeDic[AnimeType.OVA])
                    {
                        if (typeStr.Contains(str))
                        {
                            nowAnimeType = AnimeType.OVA;
                            break;
                        }
                    }
                    foreach (string str in AnimeTypeDic[AnimeType.Movie])
                    {
                        if (typeStr.Contains(str))
                        {
                            nowAnimeType = AnimeType.Movie;
                            break;
                        }
                    }
                    foreach (string str in AnimeTypeDic[AnimeType.WEB])
                    {
                        if (typeStr.Contains(str))
                        {
                            nowAnimeType = AnimeType.WEB;
                            break;
                        }
                    }
                    
                    for (int nowCol = 3; nowCol < 10; ++nowCol) 
                    {
                        if (IsBlank(sheet, nowRow, nowCol))
                        {
                            continue;
                        }
                        string titleStr = sheet.Cell(nowRow, nowCol).CachedValue.ToString();
                        foreach (string str in TitleDic[TitleType.Name])
                        {
                            if (titleStr.Contains(str))
                            {
                                nameCol = nowCol;
                                break;
                            }
                        }
                        foreach (string str in TitleDic[TitleType.OrigName])
                        {
                            if (titleStr.Contains(str))
                            {
                                origNameCol = nowCol;
                                break;
                            }
                        }
                        foreach (string str in TitleDic[TitleType.ProductionCompany])
                        {
                            if (titleStr.Contains(str))
                            {
                                productionCompanyCol = nowCol;
                                break;
                            }
                        }
                    }
                }
                if (!IsBlank(sheet, nowRow, 2) && nowAnimeType == AnimeType.TV)
                {
                    string seasonStr = sheet.Cell(nowRow, 2).CachedValue.ToString();
                    foreach (string str in SeasonDic[Season.Winter])
                    {
                        if (seasonStr.Contains(str))
                        {
                            nowSeason = Season.Winter;
                            break;
                        }
                    }
                    foreach (string str in SeasonDic[Season.Spring])
                    {
                        if (seasonStr.Contains(str))
                        {
                            nowSeason = Season.Spring;
                            break;
                        }
                    }
                    foreach (string str in SeasonDic[Season.Summer])
                    {
                        if (seasonStr.Contains(str))
                        {
                            nowSeason = Season.Summer;
                            break;
                        }
                    }
                    foreach (string str in SeasonDic[Season.Autumn])
                    {
                        if (seasonStr.Contains(str))
                        {
                            nowSeason = Season.Autumn;
                            break;
                        }
                    }
                }
                
                if (!IsBlank(sheet, nowRow, 1) || !IsBlank(sheet, nowRow, 2))
                {
                    ++nowRow;
                    continue;
                }
                
                Anime anime = new Anime();
                anime.year = year;
                anime.animeType = nowAnimeType;
                anime.season = nowSeason;
                anime.name = nameCol > 0 ? sheet.Cell(nowRow, nameCol).CachedValue.ToString() : "";
                anime.origName = origNameCol > 0 ? sheet.Cell(nowRow, origNameCol).CachedValue.ToString() : "";
                anime.productionCompany = productionCompanyCol > 0 ? sheet.Cell(nowRow, productionCompanyCol).CachedValue.ToString() : "";
                
                string statusStr = sheet.Cell(nowRow, 11).CachedValue.ToString();
                if (statusStr == "未观看")
                {
                    anime.status = Status.NeverWatched;
                }
                else if (statusStr == "正在一周目")
                {
                    anime.status = Status.Watching;
                }
                else if (statusStr == "已看过")
                {
                    anime.status = Status.Watched;
                }
                else if (statusStr == "已弃番")
                {
                    anime.status = Status.GaveUp;
                }
                string planToWatchStr = sheet.Cell(nowRow, 12).CachedValue.ToString();
                if (planToWatchStr == "是")
                {
                    anime.planToWatch = true;
                }
                else if (planToWatchStr == "否")
                {
                    anime.planToWatch = false;
                }
                anime.score = IsBlank(sheet, nowRow, 13) ? -1 : int.Parse(sheet.Cell(nowRow, 13).CachedValue.ToString());
                anime.comment = sheet.Cell(nowRow, 14).CachedValue.ToString();
                
                animeList.Add(anime);
                ++nowRow;
                
                if (anime.status == Status.Watched && anime.animeType == AnimeType.TV)
                {
                    Logger.Info(anime.year + " " + anime.season + ": " + anime.name);
                }
            }
        }

        /// <summary>
        /// すべての出力の前に呼び出されます
        /// </summary>
        /// <param name="param">着信パラメータ</param>
        /// <param name="workbook">出力用のExcelファイル</param>
        /// <param name="globalObject">グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。</param>
        /// <param name="allFilePathList">分析されたすべてのファイルパスのリスト</param>
        /// <param name="isExecuteInSequence">順番実行するかどうか</param>
        public void RunBeforeSetResult(Param param, XLWorkbook workbook, ref Object globalObject, List<string> allFilePathList, bool isExecuteInSequence)
        {
            
        }

        /// <summary>
        /// 分析結果をExcelにエクスポートする
        /// </summary>
        /// <param name="param">着信パラメータ</param>
        /// <param name="workbook">出力用のExcelファイル</param>
        /// <param name="filePath">ファイルパス</param>
        /// <param name="globalObject">グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。</param>
        /// <param name="isExecuteInSequence">順番実行するかどうか</param>
        /// <param name="invokeCount">この出力関数が呼び出された回数</param>
        /// <param name="totalCount">出力関数を呼び出す必要がある合計回数</param>
        public void SetResult(Param param, XLWorkbook workbook, string filePath, ref Object globalObject, bool isExecuteInSequence, int invokeCount, int totalCount)
        {
            
        }

        /// <summary>
        /// すべての通話が終了した後に呼び出されます
        /// </summary>
        /// <param name="param">着信パラメータ</param>
        /// <param name="workbook">出力用のExcelファイル</param>
        /// <param name="globalObject">グローバルに存在し、現在の行番号など、他の呼び出しで使用する必要のあるデータを保存できます。</param>
        /// <param name="allFilePathList">分析されたすべてのファイルパスのリスト</param>
        /// <param name="isExecuteInSequence">順番実行するかどうか</param>
        public void RunEnd(Param param, XLWorkbook workbook, ref Object globalObject, List<string> allFilePathList, bool isExecuteInSequence)
        {
            
        }
    }
}