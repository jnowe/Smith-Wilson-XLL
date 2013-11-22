using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using MathNet.Numerics.LinearAlgebra.Double;

namespace ExcelAddin
{
    public static class ExcelMaths
    {
        [ExcelFunction(Name="ADD.TWO.INTEGERS",Description = "Adds Two Integers")]
        public static int AddTwoIntegers(
            [ExcelArgument(Name = "firstNumber", Description = "This is the first integer")] int firstNumber,
            [ExcelArgument(Name = "secondNumber", Description = "This is the second integer")] int secondNumber)
        {
            int addedNumbers = firstNumber + secondNumber;
            return addedNumbers;
        }

        [ExcelFunction(Name = "GET.SMITH.WILSON.BOND.PRICE.CURVE.PARALLEL.CSHARP", Description = "Returns a bond price curve (as an array) by using the Task Parallel Library")]
        public static double[] SmithWilsonParallelCurve(
            [ExcelArgument(Name = "Inputs", Description = "1st column: bond terms excluding 0. 2nd: bond prices")] double[,] inputs,
            [ExcelArgument(Name = "UFR", Description = "Ultimate forward rate")] double ufr,
            [ExcelArgument(Name = "Alpha", Description = "Speed of mean reversion")] double alpha,
            [ExcelArgument(Name = "MonthlyCurveLength", Description = "Curve Length i.e. num rows")] int monthlyCurveLength)
        {
            var bondCurve = new double[monthlyCurveLength];
            var rows = Enumerable.Range(0, monthlyCurveLength);
            Parallel.ForEach(rows, row => 
            {
                var tau = row / 12d;
                bondCurve[row] = SmithWilson(inputs, ufr, tau, alpha);
            });
            return bondCurve;
        }

        [ExcelFunction(Name = "GET.SMITH.WILSON.BOND.PRICE.CURVE.CSHARP", Description = "Returns bond price curve as an array")]
        public static double[] SmithWilsonCurve(
            [ExcelArgument(Name = "Inputs", Description = "1st column is bond terms excluded 0. 2nd is bond prices")] double[,] inputs,
            [ExcelArgument(Name = "UFR", Description = "Ultimate forward rate")] double ufr,
            [ExcelArgument(Name = "Tau", Description = "Term of the bond being returned")] double tau,
            [ExcelArgument(Name = "Alpha", Description = "Speed of mean reversion")] double alpha)
        {
            var inputsMat = inputs;
            var length = inputsMat.GetLength(0);
            var timeVec = new double[length];
            var zcbVec = new double[length];
            var ufrVec = new double[length];
            var zetaVec = new double[length];
            var pMinusEVec = new double[length];
            var wMat = new double[length, length];
            var wInvMat = new double[length, length];
            var zetaWVec = new double[length];
            var zetaW = 0.0;
            var returnCurve = new double[1620];

            ufr = Math.Log(1 + ufr);

            for (int i = 0; i < length; i++)
            {
                timeVec[i] = inputsMat[i, 0];
                zcbVec[i] = inputsMat[i, 1];
                ufrVec[i] = Math.Exp(-ufr * timeVec[i]);
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    wMat[i, j] = DoubleYou(timeVec[i], timeVec[j], ufr, alpha);
                }
            }

            var wMatAsDenseMatrix = DenseMatrix.OfArray(wMat);
            wInvMat = wMatAsDenseMatrix.Inverse().ToArray();
            for (int i = 0; i < length; i++)
            {
                pMinusEVec[i] = zcbVec[i] - ufrVec[i];
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    zetaVec[i] = zetaVec[i] + wInvMat[i, j] * pMinusEVec[j];
                }
            }
            for (int j = 0; j < 1620; j++)
            {
                tau = (double)(j+1) / 12d;
                zetaW = 0d;
                for (int i = 0; i < length; i++)
                {
                    zetaW += zetaVec[i] * DoubleYou(tau, timeVec[i], ufr, alpha);
                }
                returnCurve[j] = Math.Exp(-ufr * tau) + zetaW;
            }

            return returnCurve;
        }
        

        [ExcelFunction(Name = "GET.SMITH.WILSON.BOND.PRICE.CSHARP", Description = "Returns a single bond price")]
        public static double SmithWilson(
            [ExcelArgument(Name = "Inputs", Description = "1st column is bond terms excluded 0. 2nd is bond prices")] double[,] inputs,
            [ExcelArgument(Name = "UFR", Description = "Ultimate forward rate")] double ufr,
            [ExcelArgument(Name = "Tau", Description = "Term of the bond being returned")] double tau,
            [ExcelArgument(Name = "Alpha", Description = "Speed of mean reversion")] double alpha)
        {
            var inputsMat = inputs;
            var length = inputsMat.GetLength(0);
            var timeVec = new double[length];
            var zcbVec = new double[length];
            var ufrVec = new double[length];
            var zetaVec = new double[length];
            var pMinusEVec = new double[length];
            var wMat = new double[length, length];
            var wInvMat = new double[length, length];
            var zetaWVec = new double[length];
            var zetaW = 0.0;

            ufr = Math.Log(1 + ufr);

            for (int i = 0; i < length; i++)
            {
                timeVec[i] = inputsMat[i, 0];
                zcbVec[i] = inputsMat[i, 1];
                ufrVec[i] = Math.Exp(-ufr * timeVec[i]);
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    wMat[i, j] = DoubleYou(timeVec[i], timeVec[j], ufr, alpha);
                }
            }

            var wMatAsDenseMatrix = DenseMatrix.OfArray(wMat);
            wInvMat = wMatAsDenseMatrix.Inverse().ToArray();
            for (int i = 0; i < length; i++)
            {
                pMinusEVec[i] = zcbVec[i] - ufrVec[i];
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    zetaVec[i] = zetaVec[i] + wInvMat[i, j] * pMinusEVec[j];
                }
            }

            for (int i = 0; i < length; i++)
            {
                zetaW += zetaVec[i] * DoubleYou(tau, timeVec[i], ufr, alpha);
            }

            return Math.Exp(-ufr * tau) + zetaW;
        }

        private static double DoubleYou(double tau, double you, double ufr, double a)
        {
            var p1 = Math.Exp(-ufr * (tau + you));
            var p2 = a * Math.Min(tau, you);
            var p3 = 0.5 * Math.Exp(-a * (Math.Max(tau, you)));
            var p4 = Math.Exp(a * Math.Min(tau, you));
            var p5 = Math.Exp(-a * Math.Min(tau, you));
            return p1 * (p2 - p3 * (p4 - p5));
        }
    }
}