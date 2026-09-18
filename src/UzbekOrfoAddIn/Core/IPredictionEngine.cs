using System.Threading;
using System.Threading.Tasks;
using UzbekOrfoAddIn.Prediction;

namespace UzbekOrfoAddIn.Core
{
    /// <summary>
    /// Predicts observed continuations from a detached collection snapshot.
    /// </summary>
    public interface IPredictionEngine
    {
        Task<PredictionCandidate[]> PredictAsync(PredictionRequest request,
            CancellationToken cancellationToken = default(CancellationToken));
    }
}
