namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AzureStorage
{
    using System.IO;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Http;
    using global::Azure.Storage.Blobs;
    using global::Azure.Storage.Blobs.Models;

    public class CDNImagesBlobContainerService
    {
        private readonly BlobContainerClient blobContainerClient;

        /// <summary>
        /// Initializes a new instance of the <see cref="CDNImagesBlobContainerService"/> class.
        /// </summary>
        /// <param name="blobContainerClientResolver">blobContainerClientResolver resolver.</param>
        public CDNImagesBlobContainerService(BlobContainerClientResolver blobContainerClientResolver)
        {
            this.blobContainerClient = blobContainerClientResolver(Common.Constants.BlobImagesCDNContainerName);
        }

        /// <summary>
        /// upload file to public cdn.
        /// </summary>
        /// <param name="file">file.</param>
        /// <param name="fileName">fileName.</param>
        /// <param name="contentType">contentType.</param>
        /// <returns>file url.</returns>
        public async Task<string> UploadFileToBlobContainer(IFormFile file, string fileName, string contentType)
        {
            BlobClient blobClient = this.blobContainerClient.GetBlobClient(fileName);
            var stream = file.OpenReadStream();
            await blobClient.UploadAsync(stream, new BlobHttpHeaders { ContentType = contentType });
            return blobClient.Uri.AbsoluteUri;
        }
    }
}
