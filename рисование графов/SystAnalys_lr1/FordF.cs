using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SystAnalys_lr1
{
    public class FordFulkerson
    {
        public FordFulkerson(int size)
        {
            NUM_VERTICES = size;
        }
        const int MAX_VERTICES = 14;

        int NUM_VERTICES; // число вершин в графе
        const int INFINITY = 10000; // условное число обозначающее бесконечность

        // f - массив содержащий текущее значение потока
        // f[i][j] - поток текущий от вершины i к j
        public int[,] f = new int[MAX_VERTICES, MAX_VERTICES];
        // с - массив содержащий вместимоти ребер,
        // т.е. c[i][j] - максимальная величину потока способная течь по ребру (i,j)
        public int[,] c = new int[MAX_VERTICES, MAX_VERTICES];
        // набор вспомогательных переменных используемых функцией FindPath - обхода в ширину
        // Flow - значение потока чарез данную вершину на данном шаге поиска
        int[] Flow = new int[MAX_VERTICES];
        // Link используется для нахождения собственно пути
        // Link[i] хранит номер предыдущей вешины на пути i -> исток
        int[] Link = new int[MAX_VERTICES];
        List<int> tmp = new List<int>();
        int[] Queue = new int[MAX_VERTICES]; // очередь
        int QP, QC; // QP - указатель начала очереди и QC - число эл-тов в очереди

        public void RecMas(int[,] a, int size)
        {
            for (int i = 0; i < size; i++)
            {
                for (int j = 0; j < size; j++)
                {
                    c[i, j] = a[i, j];
                }
            }
        }

        public int FindPath(int source, int target) // source - исток, target - сток
        {
            QP = 0; QC = 1; Queue[0] = source;
            Link[target] = -1; // особая метка для стока
            //int i;
            int CurVertex;
            //  memset(Flow, 0, sizeof(int) * NUM_VERTICES); // в начале из всех вершин кроме истока течет 0
            Array.Clear(Flow, 0, Flow.Length);
            Flow[source] = INFINITY; // а из истока может вытечь сколько угодно
            while (Link[target] == -1 && QP < QC)
            {
                // смотрим какие вершины могут быть достигнуты из начала очереди
                CurVertex = Queue[QP];
                for (int i = 0; i < NUM_VERTICES; i++)
                    // проверяем можем ли мы пустить поток по ребру (CurVertex,i):
                    if ((c[CurVertex, i] - f[CurVertex, i]) > 0 && Flow[i] == 0)
                    {

                        //  if (f[CurVertex, i] + f[Link[i - 1], CurVertex] > c[Link[i - 1], CurVertex])
                        {
                            // если можем, то добавляем i в конец очереди
                            Queue[QC] = i; QC++;
                            Link[i] = CurVertex; // указываем, что в i добрались из CurVertex
                                                 // и находим значение потока текущее через вершину i

                            if ((c[CurVertex, i] - f[CurVertex, i] < Flow[CurVertex]))
                                Flow[i] = c[CurVertex, i];
                            else
                                Flow[i] = Flow[CurVertex];
                        }

                    }
                QP++;// прерходим к следующей в очереди вершине
            }

            // закончив поиск пути
            if (Link[target] == -1) return 0; // мы или не находим путь и выходим
            // или находим:
            // тогда Flow[target] будет равен потоку который "дотек" по данному пути из истока в сток
            // тогда изменяем значения массива f для  данного пути на величину Flow[target]
            CurVertex = target;
            while (CurVertex != source) // путь из стока в исток мы восстанавливаем с помощбю массива Link
            {
                //if ((f[Link[CurVertex], CurVertex] + Flow[target]) > c[Link[CurVertex], CurVertex])
                {
                    //f[CurVertex,Link[CurVertex]] -= Flow[target];
                    // if (f[Link[CurVertex], CurVertex] + Flow[target] <= c[Link[CurVertex], CurVertex])
                    // {
                    f[Link[CurVertex], CurVertex] += Flow[target];
                    // }
                    CurVertex = Link[CurVertex];


                }
            }
            return Flow[target]; // Возвращаем значение потока которое мы еще смогли "пустить" по графу
        }




        public void DrawFlow(TextBox textBox1, Graphics graphic, PictureBox pbox, int[,] Items, List<graphic> Item)
        {

            Brush brush = new SolidBrush(Color.Red);
            Pen penelips = new Pen(brush);

            Font font = new Font("Arial", 12);
            graphic = pbox.CreateGraphics(); //-V3061
            // proverka(Items, f, Convert.ToInt16(textBox1.Text));
            for (int i = 0; i < Convert.ToInt16(textBox1.Text); i++)
            {
                for (int j = 0; j < Convert.ToInt16(textBox1.Text); j++)
                {
                    if ((Items[i, j] != 0) && (i != j) && (f[i, j]) != 0)
                    {
                        graphic.DrawString(Items[i, j].ToString() + "/" + f[i, j].ToString(), font, brush, new Point((Item[i].Locate.X + Item[j].Locate.X) / 2 - 4,
                            (Item[i].Locate.Y + Item[j].Locate.Y) / 2 - 4));

                    }
                    if (Items[i, j] != 0 && i != j)
                    {
                        graphic.DrawString(Items[i, j].ToString(), font, brush, new Point((Item[i].Locate.X + Item[j].Locate.X) / 2 - 4,
                                                  (Item[i].Locate.Y + Item[j].Locate.Y) / 2 - 4));
                    }
                }
            }

        }

        public int MaxFlow(int source, int target) // source - исток, target - сток
        {
            // инициализируем переменные:
            //memset(f, 0, sizeof(int) * MAX_VERTICES * MAX_VERTICES); // по графу ничего не течет
            Array.Clear(f, 0, f.Length);
            int MaxFlow = 0; // начальное значение потока
            int AddFlow;
            do
            {
                // каждую итерацию ищем какй-либо простой путь из истока в сток
                // и какой еще поток мажет быть пущен по этому пути
                AddFlow = FindPath(source, target - 1);
                MaxFlow += AddFlow;
            }
            while (AddFlow > 0);// повторяем цикл пока поток увеличивается

            return MaxFlow;
        }



    }
}
