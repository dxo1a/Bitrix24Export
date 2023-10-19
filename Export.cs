using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Bitrix24Export
{
    public class Export
    {
        public static readonly string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        public string FilePath, SaveFileName, MessageText, UrlText, WebhookUrl, UrlPreview;
        public int FolderID, ChatID, ChatType;

        public Export(string filePath, string saveFileName, string messageText, string urlText, string webhookUrl, string urlPreview, int folderId, int chatId, int chatType)
        {
            MessageText = messageText;
            UrlText = urlText;
            WebhookUrl = webhookUrl;
            UrlPreview = urlPreview;
            FolderID = folderId;
            ChatID = chatId;
            ChatType = chatType;

            SaveFileName = saveFileName;              // будущее имя на диске битрикса
            FilePath = filePath;                      // путь до файла на пк
        }

        public void StartExport() => ExportToBitrix24(FilePath, FolderID, ChatID, MessageText, UrlText);

        private async void ExportToBitrix24(string filePath, int folder, int chat, string messageText, string urlText)
        {
            int fileId = await UploadFileToBitrix24(filePath, folder);

            await SendMessageToChatWebhook(chat, messageText, urlText, fileId);
        }

        public async Task<int> UploadFileToBitrix24(string filePath, int folderId)
        {
            using (HttpClient client = new HttpClient())
            {
                if (File.ReadAllBytes(filePath).Length <= 0)
                {
                    Console.WriteLine("[Ошибка] Неправильно присвоен путь до файла.");
                }
                byte[] fileBytes = File.ReadAllBytes(filePath);         // используется документ который был создан при создании файла эксель
                string base64File = Convert.ToBase64String(fileBytes);

                var requestData = new
                {
                    id = folderId, // ID папки, в которую загружается файл
                    data = new
                    {
                        NAME = SaveFileName
                    },
                    fileContent = base64File,
                    generateUniqueName = true // Уникализировать имя файла, если файл с таким именем уже существует
                };

                string jsonRequestData = JsonConvert.SerializeObject(requestData);
                Console.WriteLine($"{jsonRequestData}");
                var content = new StringContent(jsonRequestData, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync($"{WebhookUrl}disk.folder.uploadfile.json", content);

                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    dynamic responseData = JsonConvert.DeserializeObject(responseBody);
                    Console.WriteLine($"Файл [{SaveFileName}] успешно сохранён в папку [{folderId}].");

                    if (responseData.result != null && responseData.result.ID != null)
                    {
                        int fileId = responseData.result.ID;
                        Console.WriteLine($"Присвоен ID: [{fileId}]");
                        return fileId;
                    }
                    else
                    {
                        throw new Exception("ID файла в ответе от сервера является null или недопустимого формата.");
                    }
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Ошибка при загрузке файла на личный диск: {response.StatusCode} - {errorContent}");
                }
            }
        }

        public async Task<string> GetFileUrlById(int fileId)
        {
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = await client.GetAsync($"{WebhookUrl}disk.file.get.json?id={fileId}");

                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    dynamic responseData = JsonConvert.DeserializeObject(responseBody);

                    if (responseData.result != null && responseData.result.DOWNLOAD_URL != null)
                    {
                        string fileUrl = responseData.result.DETAIL_URL;
                        Console.WriteLine($"Ссылка на файл с ID [{fileId}]: {fileUrl}");
                        return fileUrl;
                    }
                    else
                    {
                        throw new Exception("Невозможно получить ссылку на файл. Ответ от сервера не содержит ожидаемых данных.");
                    }
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Ошибка при получении ссылки на файл: {response.StatusCode} - {errorContent}");
                }
            }
        }

        public async Task<string> SendMessageToChatWebhook(int chatId, string message, string nameForUrl, int fileId)
        {
            string fileUrl = await GetFileUrlById(fileId);

            using (HttpClient client = new HttpClient())
            {
                var attachments = new[]
                {
                    new
                    {
                        LINK = new
                        {
                            //чтобы получить ссылку на изображение надо получить публичную в битриксе, а потом пкм->открыть изображение в новой вкладке
                            PREVIEW = UrlPreview,
                            NAME = nameForUrl,
                            LINK = fileUrl,
                        }
                    }
                };

                string jsonMessageData;
                if (ChatType == 0)
                {
                    var messageData = new
                    {
                        DIALOG_ID = chatId.ToString(),
                        MESSAGE = message,
                        ATTACH = attachments
                    };
                    jsonMessageData = JsonConvert.SerializeObject(messageData);
                }
                else
                {
                    var messageData = new
                    {
                        CHAT_ID = chatId.ToString(),
                        MESSAGE = message,
                        ATTACH = attachments
                    };
                    jsonMessageData = JsonConvert.SerializeObject(messageData);
                }

                var content = new StringContent(jsonMessageData, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync($"{WebhookUrl}im.message.add.json", content);

                if (response.IsSuccessStatusCode)
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    return responseBody;
                }
                else
                {
                    string errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Ошибка при отправке сообщения в чат: {response.StatusCode} - {errorContent}");
                }
            }
        }
    }
}
