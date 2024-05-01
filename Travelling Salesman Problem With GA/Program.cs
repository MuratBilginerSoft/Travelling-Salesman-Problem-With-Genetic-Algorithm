using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Travelling_Salesman_Problem_With_GA
{
    #region Classes
    public class Locations
    {
        public string point;
        public int x;
        public int y;
    }

    #endregion

    internal class Program
    {
        
        #region Functions

        static double CalcDistance(int x1, int x2, int y1, int y2)
        {
            double result = Math.Sqrt(Math.Pow((x2 - x1), 2) + Math.Pow((y2 - y1), 2));
            return result;
        }

        static void ReadExcel(string path, Application excelApp, List<Locations> locationList)
        {
            Workbook workBook = excelApp.Workbooks.Open(path);
            Worksheet workSheet = workBook.Worksheets[1];

            for (int i = 1; i < 9; i++)
            {
                Locations l1 = new Locations();

                l1.point = workSheet.Cells[i, 1].Value;
                l1.x = Convert.ToInt32(workSheet.Cells[i, 2].Value);
                l1.y = Convert.ToInt32(workSheet.Cells[i, 3].Value);

                locationList.Add(l1);
            }

            workBook.Close(0);
            excelApp.Quit();
        }

        static void SelectPopulation(Random random,string[] points, List<List<string>> population, List<List<string>> allPopulation)
        {
            for (int i = 0; i < 10; i++)
            {
                List<string> selectPoint = new List<string>();
                selectPoint.Add("Depo");

                for (int j = 1; j < 8; j++)
                {
                    bool selectStatus = true;

                    while (selectStatus)
                    {
                        int randNumb = random.Next(0, 7);
                        bool pointStatus = selectPoint.Contains(points[randNumb]);

                        if (pointStatus == false)
                        {
                            selectPoint.Add(points[randNumb]);
                            selectStatus = false;
                        }
                    }

                }

                selectPoint.Add("Depo");
                population.Add(selectPoint);
                allPopulation.Add(selectPoint);

               
            }

            Console.WriteLine("***** İlk Populasyon *********");

            foreach (var item in allPopulation)
            {
                for (int k = 0; k < item.Count; k++)
                {
                    Console.Write(item[k] + " - ");
                }

                Console.WriteLine();
            }
        }

        static void CalcGeneDistance(List<List<string>> population, List<Locations> locationList, List<double> allDistance, double distanceOne, double sumDistance, double minDistance, List<List<string>> bestGene, List<double> bestDistance, int controlA)
        {
            if (controlA == 0)
            {
                controlA = 0;

                foreach (var item in population)
                {
                    controlA++;
                    sumDistance = 0;

                    for (int i = 0; i < item.Count; i++)
                    {
                        if (i < item.Count - 1)
                        {
                            List<Locations> l1 = locationList.Where(x => x.point == item[i]).ToList();
                            List<Locations> l2 = locationList.Where(x => x.point == item[i + 1]).ToList();

                            distanceOne = CalcDistance(l1[0].x, l2[0].x, l1[0].y, l2[0].y);
                            sumDistance += distanceOne;
                        }
                    }

                    allDistance.Add(sumDistance);

                    if (controlA == 1)
                    {
                        minDistance = sumDistance;
                        bestGene.Add(item);
                        bestDistance.Add(minDistance);
                    }
                    else
                    {
                        if (sumDistance <= minDistance)
                        {
                            minDistance = sumDistance;
                            bestGene.Add(item);
                            bestDistance.Add(minDistance);
                        }
                    }
                }

                for (int k = 0; k < allDistance.Count; k++)
                {
                    Console.WriteLine(allDistance[k]);

                }

                Console.WriteLine("******** İlk En İyi Mesafeler *********");

                int o1 = 0;
                foreach (var item in bestGene)
                {
                    for (int k = 0; k < item.Count; k++)
                    {
                        Console.Write(item[k] + " - ");
                    }

                    Console.Write(bestDistance[o1]);
                    o1++;

                    Console.WriteLine();
                }
            }

            else
            {
                foreach (var item in population)
                {
                    sumDistance = 0;

                    for (int i = 0; i < item.Count; i++)
                    {
                        if (i < item.Count - 1)
                        {
                            List<Locations> l1 = locationList.Where(x => x.point == item[i]).ToList();
                            List<Locations> l2 = locationList.Where(x => x.point == item[i + 1]).ToList();

                            distanceOne = CalcDistance(l1[0].x, l2[0].x, l1[0].y, l2[0].y);

                            sumDistance += distanceOne;
                        }
                    }

                    allDistance.Add(sumDistance);

                    if (sumDistance <= minDistance)
                    {
                        minDistance = sumDistance;
                        bestGene.Add(item);
                        bestDistance.Add(minDistance);
                    }

                }

                for (int k = 0; k < allDistance.Count; k++)
                {
                    Console.WriteLine(allDistance[k]);

                }

                Console.WriteLine("******** Tüm En İyi Mesafeler *********");

                int o1 = 0;
                foreach (var item in bestGene)
                {
                    for (int k = 0; k < item.Count; k++)
                    {
                        Console.Write(item[k] + " - ");
                    }

                    Console.Write(bestDistance[o1]);
                    o1++;

                    Console.WriteLine();
                }
            }
            
        }
        
        static void CreateNewGeneration(List<List<string>> population, bool genStatus, List<List<string>> allPopulation)
        {
            #region Create New Generation Loop

            for (int i = 0; i < 10; i = i + 2)
            {
                if (i < 9)
                {
                    #region Select Gene

                    List<string> g1 = population[i];
                    List<string> g2 = population[i+1];

                    List<string> ng1 = new List<string>();
                    List<string> ng2 = new List<string>();

                    #endregion

                    #region New Gene 1

                    for (int j = 0; j < 5; j++)
                    {
                        ng1.Add(g1[j]);
                    }

                    for (int l = 5; l < 8; l++)
                    {
                        genStatus = ng1.Contains(g2[l]);

                        if (genStatus == false)
                        {
                            ng1.Add(g2[l]);
                        }
                    }

                    for (int k = 0; k < 8; k++)
                    {
                        genStatus = ng1.Contains(g1[k]);

                        if (genStatus == false)
                        {
                            ng1.Add(g1[k]);
                        }
                    }

                    ng1.Add("Depo");

                    population.Add(ng1);
                    allPopulation.Add(ng1);


                    #endregion

                    #region New Gene 2

                    for (int j = 0; j < 5; j++)
                    {
                        ng2.Add(g2[j]);
                    }

                    for (int l = 5; l < 8; l++)
                    {
                        genStatus = ng2.Contains(g1[l]);

                        if (genStatus == false)
                        {
                            ng2.Add(g1[l]);
                        }
                    }

                    for (int k = 0; k < 8; k++)
                    {
                        genStatus = ng2.Contains(g2[k]);

                        if (genStatus == false)
                        {
                            ng2.Add(g2[k]);
                        }
                    }

                    ng2.Add("Depo");

                    population.Add(ng2);
                    allPopulation.Add(ng2);

                    #endregion
                }
            }

            #endregion

            #region Remove Old Population Gene 

            for (int i = 0; i < 10; i++)
            {
                population.RemoveAt(i);
            }

            #endregion
        }

        static void Mutation(Random random,int r1, int r2, List<List<string>> population)
        {
            for (int i = 0; i < 10; i++)
            {
                r1 = random.Next(1, 8);

                bool status2 = true;

                do
                {
                    r2 = random.Next(1, 8);

                    if (r1 == r2)
                    {
                        status2 = true;
                    }
                    else
                    {
                        status2 = false;
                    }

                } while (status2);

                string gen1 = population[i][r1];
                string gen2 = population[i][r2];

                population[i][r1] = gen2;
                population[i][r2] = gen1;
            }
        }

        static void BetterMutation(Random random, int r1, int r2, List<string> betterGene, List<List<string>> allPopulation, double sumDistance, double minDistance, double distanceOne, List<Locations> locationList, List<double> allDistance, List<List<string>> bestGene, List<double> bestDistance)
        {
            r1 = random.Next(1, 8);

            bool status2 = true;

            do
            {
                r2 = random.Next(1, 8);

                if (r1 == r2)
                {
                    status2 = true;
                }
                else
                {
                    status2 = false;
                }

            } while (status2);

            string gen1 = betterGene[r1];
            string gen2 = betterGene[r2];

            betterGene[r1] = gen2;
            betterGene[r2] = gen1;
            allPopulation.Add(betterGene);

            sumDistance = 0;

            for (int i = 0; i < betterGene.Count; i++)
            {
                if (i < betterGene.Count - 1)
                {
                    List<Locations> l1 = locationList.Where(y => y.point == betterGene[i]).ToList();
                    List<Locations> l2 = locationList.Where(y => y.point == betterGene[i + 1]).ToList();

                    distanceOne = CalcDistance(l1[0].x, l2[0].x, l1[0].y, l2[0].y);

                    sumDistance += distanceOne;
                }
            }

            allDistance.Add(sumDistance);

            if (sumDistance < minDistance)
            {
                minDistance = sumDistance;
                bestGene.Add(betterGene);
                bestDistance.Add(minDistance);
            }
        }


        static void Result(List<List<string>> bestGene, List<List<string>> allPopulation, List<double> bestDistance, List<double> allDistance, int controlB)
        {
            foreach (var item in bestGene)
            {
                for (int i = 0; i < item.Count; i++)
                {
                    Console.Write(item[i] + " - ");
                }

                Console.Write(bestDistance[controlB]);
                Console.WriteLine();
                controlB++;
            }

            Console.WriteLine("********** Tüm Populasyon Üyeleri ***********");

            controlB = 0;

            foreach (var item in allPopulation)
            {
                for (int i = 0; i < item.Count; i++)
                {
                    Console.Write(item[i] + " - ");
                }

                Console.Write(allDistance[controlB]);
                Console.WriteLine();
                controlB++;
            }
        }
        #endregion



        static void Main(string[] args)
        {

            #region TODO

            #region TODO 1 Locations Sınıfını Oluştur

            /* Bu sınıf dosyadan okunan verinin Lokasyondaki ismini point ile X ve Y konumundaki değerlerini x ve y değerlerinde tutacak. */

            #endregion

            #region TODO 2 Datayı Oku

            /* Data Excel dosyasından okunup LocationList sınıfından üretilen nesneye atanacak */

            #endregion

            #region TODO 3 Populasyonu Oluştur

            /* Başlangıç populasyonu 6 değerli olarak belirlendi ve rassal olarak benzersiz 6 tane rota oluşturuldu */

            #endregion

            #region TODO 4 Populasyonun Üyelerinin Mesafelerini Hesapla

            /* Populasyon içindeki her bir elemanın toplam oluşturduğu mesafe değeri var bunlar her üye için yapılıp bir listede tutulacak. */

            #endregion

            #region TODO 5 Çaprazlama İşlemini Gerçekleştir

            /* Çaprazlama işlemi sıralı tek noktalı ve %50 oran ile gerçekleştirilecektir. Yani Populasyon içindeki ebeveynler 1 den başlayarak 2 şerli olarak ele alınacak ve iki ebeveyn gen ortadan bölünerek birleştirilecektir. Böylelikle yeni genler oluşacaktır.  */

            #endregion

            #region TODO 6 Mutasyon İşlemini Gerçekleştir

            /* 
             Mutasyon işlemi rassal olarak seçilen 2 genin yer değiştirmesi şeklinde gerçekleştirilir.
             */

            #endregion

            #region TODO 7 En İyi Sonuçları Yazdır

            // Tüm çaprazlama ve mutasyon işlemleri sonucunda her iteraston adımında gerçekleşmiş en iyi değerleri sakla ve ekrana yazdır

            #endregion

            #endregion

            #region Create Object

            Random random = new Random();
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            List<Locations> locationList = new List<Locations>();
            List<List<string>> population = new List<List<string>>();
            List<List<string>> allPopulation = new List<List<string>>();
            List<List<string>> bestGene = new List<List<string>>();
            List<double> bestDistance = new List<double>();
            List<double> allDistance = new List<double>();

            #endregion

            #region Variables

            string x = System.Windows.Forms.Application.StartupPath;
            string path = x +  @"\Data\Locations.xlsx"; // Excel Dosyasının Yolu

            string[] points = new string[] { "A", "B", "C", "D", "E", "F", "G" };

            bool genStatus = false;

            double distanceOne = 0;
            double sumDistance = 0;
            double minDistance = 0;

            #endregion

            #region Read Data

            ReadExcel(path, excelApp, locationList);

            #endregion

            #region Select Population

            SelectPopulation(random, points, population,allPopulation);

            #endregion

            #region Calculate the Distances of the Members of the First Population

            CalcGeneDistance(population, locationList, allDistance, distanceOne, sumDistance, minDistance, bestGene, bestDistance, 0);

            #endregion

            #region İteration Process

            for (int z = 0; z < 10; z++)
            {
                #region Create New Generation

                CreateNewGeneration(population, genStatus, allPopulation);

                #endregion


                #region Mutation Process

                Mutation(random, 0, 0, population);

                #endregion

                #region Calculate the Distances of the Members of the Population

                CalcGeneDistance(population, locationList,allDistance, distanceOne, sumDistance, minDistance, bestGene, bestDistance, -1);

                #endregion
            }

            #endregion

            #region Result

            int count = bestGene.Count;
            Console.WriteLine(count);

            List<string> betterGene = new List<string>();
            betterGene = bestGene[count - 1];

            for (int j = 0; j < 100; j++)
            {
                BetterMutation(random, 0, 0, betterGene,allPopulation,sumDistance,minDistance,distanceOne,locationList,allDistance,bestGene,bestDistance);   
            }

            //Result(bestGene, allPopulation, bestDistance, allDistance, 0);

            Console.ReadLine();

            #endregion

        }
    }
}
