using System;
using System.Collections.Generic;

namespace Common.Excel
{
    public interface IStatusProgress : IAggregateProgress
    {
        void Report(string status, double percents);
    }

    public interface IAggregateProgress : IProgress<double>
    {
        IReadOnlyList<IAggregateProgress> CreateParallelProgresses(IReadOnlyCollection<double> weightes);
        IReadOnlyList<IAggregateProgress> CreateParallelProgresses(int count);
        IReadOnlyList<IAggregateProgress> CreateParallelProgresses(params double[] wieghtes);
    }
}
