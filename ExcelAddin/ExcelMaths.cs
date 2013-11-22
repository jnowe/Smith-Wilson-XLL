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

        [ExcelFunction(Name = "GET.SMITH.WILSON.BOND.PRICE.CURVE.PARALLEL.CSHARP", Description = "Returns bond price curve by using the Task Parallel Library")]
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

        [ExcelFunction(Name = "GET.SMITH.WILSON.BOND.PRICE.CURVE.CSHARP", Description = "Returns bond price curve")]
        public static double[] SmithWilsonCurve(
            [ExcelArgument(Name = "Inputs", Description = "1st column is bond terms excluded 0. 2nd is bond prices")] double[,] Inputs,
            [ExcelArgument(Name = "UFR", Description = "Ultimate forward rate")] double UFR,
            [ExcelArgument(Name = "Tau", Description = "Term of the bond being returned")] double Tau,
            [ExcelArgument(Name = "Alpha", Description = "Speed of mean reversion")] double a)
        {
            var InputsMat = Inputs;
            var length = InputsMat.GetLength(0);
            var TimeVec = new double[length];
            var ZCBVec = new double[length];
            var UFRVec = new double[length];
            var ZetaVec = new double[length];
            var PminusEVec = new double[length];
            var WMat = new double[length, length];
            var WinvMat = new double[length, length];
            var ZetaWVec = new double[length];
            var ZetaW = 0.0;
            var ReturnCurve = new double[1620];

            UFR = Math.Log(1 + UFR);

            for (int i = 0; i < length; i++)
            {
                TimeVec[i] = InputsMat[i, 0];
                ZCBVec[i] = InputsMat[i, 1];
                UFRVec[i] = Math.Exp(-UFR * TimeVec[i]);
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    WMat[i, j] = DoubleYou(TimeVec[i], TimeVec[j], UFR, a);
                }
            }

            var WMatAsDenseMatrix = DenseMatrix.OfArray(WMat);
            WinvMat = WMatAsDenseMatrix.Inverse().ToArray();
            for (int i = 0; i < length; i++)
            {
                PminusEVec[i] = ZCBVec[i] - UFRVec[i];
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    ZetaVec[i] = ZetaVec[i] + WinvMat[i, j] * PminusEVec[j];
                }
            }
            for (int j = 0; j < 1620; j++)
            {
                Tau = (double)(j+1) / 12d;
                ZetaW = 0d;
                for (int i = 0; i < length; i++)
                {
                    ZetaW += ZetaVec[i] * DoubleYou(Tau, TimeVec[i], UFR, a);
                }
                ReturnCurve[j] = Math.Exp(-UFR * Tau) + ZetaW;
            }

            return ReturnCurve;
        }
        

        [ExcelFunction(Name = "GET.SMITH.WILSON.BOND.PRICE.CSHARP", Description = "Returns a single bond price")]
        public static double SmithWilson(
            [ExcelArgument(Name = "Inputs", Description = "1st column is bond terms excluded 0. 2nd is bond prices")] double[,] Inputs,
            [ExcelArgument(Name = "UFR", Description = "Ultimate forward rate")] double UFR,
            [ExcelArgument(Name = "Tau", Description = "Term of the bond being returned")] double Tau,
            [ExcelArgument(Name = "Alpha", Description = "Speed of mean reversion")] double a)
        {
            var InputsMat = Inputs;
            var length = InputsMat.GetLength(0);
            var TimeVec = new double[length];
            var ZCBVec = new double[length];
            var UFRVec = new double[length];
            var ZetaVec = new double[length];
            var PminusEVec = new double[length];
            var WMat = new double[length, length];
            var WinvMat = new double[length, length];
            var ZetaWVec = new double[length];
            var ZetaW = 0.0;

            UFR = Math.Log(1 + UFR);

            for (int i = 0; i < length; i++)
            {
                TimeVec[i] = InputsMat[i, 0];
                ZCBVec[i] = InputsMat[i, 1];
                UFRVec[i] = Math.Exp(-UFR * TimeVec[i]);
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    WMat[i, j] = DoubleYou(TimeVec[i], TimeVec[j], UFR, a);
                }
            }

            var WMatAsDenseMatrix = DenseMatrix.OfArray(WMat);
            WinvMat = WMatAsDenseMatrix.Inverse().ToArray();
            for (int i = 0; i < length; i++)
            {
                PminusEVec[i] = ZCBVec[i] - UFRVec[i];
            }

            for (int i = 0; i < length; i++)
            {
                for (int j = 0; j < length; j++)
                {
                    ZetaVec[i] = ZetaVec[i] + WinvMat[i, j] * PminusEVec[j];
                }
            }

            for (int i = 0; i < length; i++)
            {
                ZetaW += ZetaVec[i] * DoubleYou(Tau, TimeVec[i], UFR, a);
            }

            return Math.Exp(-UFR * Tau) + ZetaW;
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