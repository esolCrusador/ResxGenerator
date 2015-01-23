using System;
using System.Collections.Generic;
using System.Linq;
using Common.Excel;

namespace ResxPackage.Dialog.Models
{
    public class StatusProgress : Progress, IStatusProgress
    {
        private readonly Action<double> _onProgress;
        private readonly Action<string, double> _onStatusProgress;

        public StatusProgress(Action<double> onProgress, Action<string, double> onStatusProgress)
            :base(null, 0)
        {
            _onProgress = onProgress;
            _onStatusProgress = onStatusProgress;
        }

        public override void Report(double value)
        {
            ProgressValue = value;
            _onProgress(value);
        }

        protected override void Increment(double progressDiff)
        {
            ProgressValue += progressDiff;
            Report(ProgressValue);
        }

        public void Report(string status, double percents)
        {
            ProgressValue = percents;
            _onStatusProgress(status, percents);
        }

        public void Clear()
        {
            Report(string.Empty, 0);
        }
    }

    public class Progress : IAggregateProgress
    {
        private readonly Progress _parentProgress;
        private readonly double _weight;
        protected double ProgressValue;

        public Progress(Progress parentProgress, double weight)
        {
            _parentProgress = parentProgress;
            _weight = weight;
        }

        protected virtual void Increment(double progressDiff)
        {
            ProgressValue += progressDiff;

            _parentProgress.Increment(_weight * progressDiff);
        }

        public IReadOnlyList<IAggregateProgress> CreateParallelProgresses(IReadOnlyCollection<double> weightes)
        {
            double totalWeight = weightes.Sum();

            return weightes.Select(w => new Progress(this, w/totalWeight)).ToList();
        }

        public IReadOnlyList<IAggregateProgress> CreateParallelProgresses(int count)
        {
            return CreateParallelProgresses(Enumerable.Repeat((double)1, count).ToList());
        }

        public IReadOnlyList<IAggregateProgress> CreateParallelProgresses(params double[] wieghtes)
        {
            return CreateParallelProgresses((IReadOnlyCollection<double>) wieghtes);
        }

        public virtual void Report(double value)
        {
            var difference = value - ProgressValue;
            ProgressValue = value;

            _parentProgress.Increment(_weight * difference);
        }
    }
}
