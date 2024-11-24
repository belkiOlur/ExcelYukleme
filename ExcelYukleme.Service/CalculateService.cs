﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelYukleme.Service
{
    public class CalculateService : ICalculateService
    {
        public CalculateService() { }
        public double CalculateSimilarity(string source, string target)
        {
            int distance = LevenshteinDistance(source, target);
            int maxLength = Math.Max(source.Length, target.Length);

            return 1.0 - (double)distance / maxLength;
        }

        private static int LevenshteinDistance(string source, string target)
        {
            int[,] matrix = new int[source.Length + 1, target.Length + 1];

            for (int i = 0; i <= source.Length; i++) matrix[i, 0] = i;
            for (int j = 0; j <= target.Length; j++) matrix[0, j] = j;

            for (int i = 1; i <= source.Length; i++)
            {
                for (int j = 1; j <= target.Length; j++)
                {
                    int cost = (source[i - 1] == target[j - 1]) ? 0 : 1;

                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost
                    );
                }
            }

            return matrix[source.Length, target.Length];
        }
    }
}
