namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.AzureStorage
{
    using global::Azure.Storage.Blobs;

    public delegate BlobContainerClient BlobContainerClientResolver(string containerName);
}
