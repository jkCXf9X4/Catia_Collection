using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Add_bodies
{
    public class ProgressBar
    {
        int max;
        int intervalls;
        double steps;
        int roundedStep;

        int counter = -1;


        public ProgressBar(int max, int intervalls)
        {
            this.max = max;
            this.intervalls = intervalls;

            this.steps = max / intervalls;
            this.roundedStep = (int)Math.Round((double)this.steps, 0);
        }


        public void Update(int? index)
        {
            int localIndex;
            if (index == null)
            {
                localIndex = this.counter++;
            }
            else
            {
                localIndex = index.Value;
            }

            if (localIndex == 0)
            {
                Console.WriteLine();
            }
            else 

            if ( localIndex % roundedStep == 0)
            {
                var current = Math.Round(localIndex / this.steps, 1);
                Console.SetCursorPosition(0, Console.CursorTop);
                Console.Write($"{current * 10}%                                              ");
            }

            if (localIndex == max - 2)
            {
                Console.WriteLine("\n");
                return;
            }
        }
    }
}
