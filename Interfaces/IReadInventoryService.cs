using System.Threading.Tasks;

namespace AshfordSync.Interfaces
{
    public interface IReadInventoryService
    {
        Task ReadInventoryAsync(int supplierId, string fileName);
        string SkuTranslate(string sku);
        Task ReadRMAAsync(int supplierId, string fileName);
        Task ReadShipConfirmAsync(int supplierId, string fileName);
        Task ReadJsonShipConfirmAsync(int supplierId, string fileName);

    }
}
