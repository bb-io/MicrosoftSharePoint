using Blackbird.Applications.Sdk.Common;
using Blackbird.Applications.Sdk.Common.Dynamic;
using Blackbird.Applications.Sdk.Common.Invocation;

namespace Apps.MicrosoftSharePoint.DataSourceHandlers;

public class ContentTypeDataSourceHandler : BaseInvocable, IDataSourceHandler
{
    public ContentTypeDataSourceHandler(InvocationContext invocationContext) : base(invocationContext)
    {
    }

    public Dictionary<string, string> GetData(DataSourceContext context)
    {
        var contentTypes = new Dictionary<string, string>
        {
            { "audio/aac", ".aac - AAC audio" },
            { "application/x-abiword", ".abw - AbiWord document" },
            { "application/x-freearc", ".arc - Archive document (multiple files embedded)" },
            { "image/avif", ".avif - AVIF image" },
            { "video/x-msvideo", ".avi - AVI: Audio Video Interleave" },
            { "application/vnd.amazon.ebook", ".azw - Amazon Kindle eBook format" },
            { "image/bmp", ".bmp - Windows OS/2 Bitmap Graphics" },
            { "application/x-bzip", ".bz - BZip archive" },
            { "application/x-bzip2", ".bz2 - BZip2 archive" },
            { "application/x-cdf", ".cda - CD audio" },
            { "application/x-csh", ".csh - C-Shell script" },
            { "text/css", ".css - Cascading Style Sheets (CSS)" },
            { "text/csv", ".csv - Comma-separated values (CSV)" },
            { "application/msword", ".doc - Microsoft Word" },
            { "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx - Microsoft Word (OpenXML)" },
            { "application/epub+zip", ".epub - Electronic publication (EPUB)" },
            { "application/gzip", ".gz - GZip Compressed Archive" },
            { "image/gif", ".gif - Graphics Interchange Format (GIF)" },
            { "text/html", ".htm/.html - HyperText Markup Language (HTML)" },
            { "image/vnd.microsoft.icon", ".ico - Icon format" },
            { "image/jpeg", ".jpeg/.jpg - JPEG images" },
            { "application/json", ".json - JSON format" },
            { "text/javascript", ".js - JavaScript" },
            { "audio/mpeg", ".mp3 - MP3 audio" },
            { "video/mp4", ".mp4 - MP4 video" },
            { "video/mpeg", ".mpeg - MPEG Video" },
            { "application/vnd.oasis.opendocument.presentation", ".odp - OpenDocument presentation document" },
            { "application/vnd.oasis.opendocument.spreadsheet", ".ods - OpenDocument spreadsheet document" },
            { "application/vnd.oasis.opendocument.text", ".odt - OpenDocument text document" },
            { "audio/ogg", ".oga - OGG audio" },
            { "video/ogg", ".ogv - OGG video" },
            { "application/ogg", ".ogx - OGG" },
            { "audio/opus", ".opus - Opus audio" },
            { "image/png", ".png - Portable Network Graphics" },
            { "application/pdf", ".pdf - Adobe Portable Document Format (PDF)" },
            { "application/vnd.ms-powerpoint", ".ppt - Microsoft PowerPoint" },
            { "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx - Microsoft PowerPoint (OpenXML)" },
            { "application/vnd.rar", ".rar - RAR archive" },
            { "application/rtf", ".rtf - Rich Text Format (RTF)" },
            { "image/svg+xml", ".svg - Scalable Vector Graphics (SVG)" },
            { "application/x-tar", ".tar - Tape Archive (TAR)" },
            { "image/tiff", ".tif/.tiff	- Tagged Image File Format (TIFF)" },
            { "text/plain", ".txt - Text" },
            { "application/vnd.visio", ".vsd, Microsoft Visio" },
            { "audio/wav", ".wav - Waveform Audio Format" },
            { "audio/webm", ".weba - WEBM audio" },
            { "video/webm", ".webm - WEBM video" },
            { "image/webp", ".webp - WEBP image" },
            { "application/xhtml+xml", ".xhtml - XHTML" },
            { "application/vnd.ms-excel", ".xls - Microsoft Excel" },
            { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx - Microsoft Excel (OpenXML)" },
            { "application/xml", ".xml - XML" },
            { "application/zip", ".zip - ZIP archive" },
            { "application/x-7z-compressed", ".7z - 7-zip archive" },
        };
        return contentTypes.Where(contentType => context.SearchString == null 
                                                 || contentType.Key.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase)
                                                 || contentType.Value.Contains(context.SearchString, StringComparison.OrdinalIgnoreCase))
            .ToDictionary(contentType => contentType.Key, contentType => contentType.Value);
    }
}