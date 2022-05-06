using MimeKit;
using Zue.Common;

namespace MailKitExtensions.Attachments
{
    public static class AttachmentExtensions
    {
        private static ILog _logger =
            LogUtil.GetLogger(nameof(AttachmentExtensions));

        #region File Handling
        public static async Task<Stream> GetFileStreamAsync(
            string filePath, CancellationToken ct = default)
        {
            const int BufferSize = 8192;
            var outputStream = new MemoryStream();
            if (FileHandler.FileCheckOk(filePath, true))
            {
                using var source = new FileStream(
                    filePath, FileMode.Open, FileAccess.Read,
                    FileShare.Read, BufferSize, useAsync: true);
                await source.CopyToAsync(outputStream, ct);
                outputStream.Position = 0;
                _logger.Debug("OK '{0}'.", filePath);
            }
            return outputStream;
        }

        public static MimeEntity? GetMimeEntityFromFilePath(string filePath)
        {
            MimeEntity? result = null;
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                var stream = GetFileStreamAsync(filePath)
                    .GetAwaiter().GetResult();
                string fileName = Path.GetFileName(filePath);
                string contentType = Path.GetExtension(fileName)
                    .Equals(".pdf", StringComparison.OrdinalIgnoreCase) ?
                        System.Net.Mime.MediaTypeNames.Application.Pdf :
                        System.Net.Mime.MediaTypeNames.Application.Octet;
                result = stream.GetMimePart(fileName, contentType);
            }
            return result;
        }

        public static MimeEntity? GetMimePart(
            this Stream stream, string fileName, string contentType = "", string contentId = "")
        {
            MimeEntity? result = null;
            if (stream != null && stream.Length > 0)
            {
                stream.Position = 0; // reset stream position ready to read
                if (string.IsNullOrWhiteSpace(contentType))
                    contentType = System.Net.Mime.MediaTypeNames.Application.Octet;
                if (string.IsNullOrWhiteSpace(contentId))
                    contentId = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                //streamIn.CopyTo(streamOut, 8192);
                result = new MimePart(contentType)
                {
                    Content = new MimeContent(stream),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentId = contentId,
                    FileName = fileName
                };
            }
            return result;
        }

        public static IEnumerable<MimeEntity> GetMimeEntitiesFromFilePaths(this IEnumerable<string> filePaths)
        {
            if (filePaths != null && filePaths.Count() == 1)
            {
                string firstFilePath = filePaths.First();
                if (!string.IsNullOrWhiteSpace(firstFilePath))
                    filePaths = firstFilePath.Split('|', StringSplitOptions.RemoveEmptyEntries); // ';' is a valid file name character
            }
            return filePaths?.Select(name => GetMimeEntityFromFilePath(name))
                ?.Where(att => att != null).ToArray() as IList<MimeEntity> ?? Array.Empty<MimeEntity>();
        }
        #endregion

        #region Mime Entity
        public static IEnumerable<string> GetAttachmentNames(
            this IEnumerable<MimeEntity> mimeEntities)
        {
            return mimeEntities?.Select(a => a.GetAttachmentName()) ?? Array.Empty<string>();
        }

        public static string GetAttachmentName(this MimeEntity mimeEntity)
        {
            string fileName = string.Empty;
            if (mimeEntity is MimePart mimePart)
                fileName = mimePart.FileName;
            else if (mimeEntity is MessagePart msgPart)
                fileName = msgPart.Message?.MessageId ??
                    msgPart.Message?.References?.FirstOrDefault() ??
                    msgPart.GetHashCode() + ".eml";
            else if (mimeEntity is MimeKit.Tnef.TnefPart tnefPart)
                fileName = tnefPart.ExtractAttachments()
                    .Select(t => t?.ContentDisposition?.FileName)
                    .ToEnumeratedString(".tnef, ");
            else
                fileName = mimeEntity?.ContentDisposition?.FileName ?? "";
            return fileName ?? "";
        }

        public static IEnumerable<MimeEntity> GetFilteredAttachments(
            this IEnumerable<MimeEntity> mimeEntities, IEnumerable<string> mediaTypes)
        {
            return mediaTypes == null || mimeEntities == null ? mimeEntities : mimeEntities
                ?.Where(a => a.IsAttachment && a is MimePart att && mediaTypes.Any(s =>
                    att.FileName?.EndsWith(s, StringComparison.OrdinalIgnoreCase) ?? false))
                 ?? Array.Empty<MimeEntity>();
        }

        public static async Task<MemoryStream> GetMimeEntityStream(
            this MimeEntity mimeEntity, CancellationToken ct = default)
        {
            var memoryStream = new MemoryStream();
            if (mimeEntity != null)
                await mimeEntity.WriteToStreamAsync(memoryStream, ct);
            memoryStream.Position = 0;
            return memoryStream;
        }

        public static async Task<Stream> WriteToStreamAsync(
            this MimeEntity entity, Stream stream, CancellationToken ct = default)
        {
            if (entity is MessagePart messagePart)
            {
                await messagePart.Message.WriteToAsync(stream, ct);
            }
            else if (entity is MimePart mimePart && mimePart.Content != null)
            {
                await mimePart.Content.DecodeToAsync(stream, ct);
            }
            // rewind the stream so the next process can read it from the beginning
            stream.Position = 0;
            return stream;
        }
        #endregion
    }
}
