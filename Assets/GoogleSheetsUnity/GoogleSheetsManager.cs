// System
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

// Google APIS
using Google;
using Google.Apis.Requests;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;

// LightJson
using LightJson;

// Unity 
using UnityEngine;

public delegate void ServiceInitializedResponse(bool success);
public delegate void SheetsRequestResponse(bool success, JsonObject sheets);

public class GoogleSheetsManager : MonoBehaviour
{
    private static string applicationName = "Unity";

    public string apiKey = "AIzaSyAbY5-EjROctQUECA8RWai9q_c8Uzlusvc";
    public string spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
    public bool immediatelyRequestData = false;

    public bool initialized { get; private set; } = false;
    public bool retrievingData { get; private set; } = false;

    private SheetsRequestResponse onRequestResponse = null;

    private JsonObject sheets = null;

    private SheetsService service = null;

    private string fullCredentialPath = "";
    private string fullTokenPath = "";

    private void Awake()
    {
        InitializeServiceApiKey();
    }

    private void Update()
    {
        if(initialized && immediatelyRequestData)
        {
            immediatelyRequestData = false;
            GetSheets(null);
        }
    }

    private bool InitializeServiceApiKey()
    {
        if (!initialized)
        {
            service = new SheetsService(new BaseClientService.Initializer()
            {
                ApiKey = apiKey,
                ApplicationName = applicationName,
            });

            initialized = true;
            return true;
        }               
        return false;
    }

    public async void GetSheets(SheetsRequestResponse response)
    {
        onRequestResponse += response;
        if (initialized && !retrievingData)
        {
            bool success = false;
            retrievingData = true;                      
            try
            {
                // Need to identify whether or not an exception / what exception is thrown for http errors
                sheets = await GetSheetsTask();
                // Lazy, assumes call was successful if no exceptions are thrown
                success = true;
            }
            catch (AggregateException ae)
            {
                ae.Handle((Exception e) =>
                {
                    if (e is GoogleApiException)
                    {
                        HandleApiException(e as GoogleApiException);
                    }
                    else
                    {
                        HandleException(e);
                    }
                    return true;
                });
            }
            catch (GoogleApiException e)
            {
                HandleApiException(e);
            }
            catch (Exception e)
            {
                HandleException(e);
            }
            retrievingData = false;

            onRequestResponse?.Invoke(success, sheets);
            onRequestResponse = null;
        }
    }

    private async Task<JsonObject> GetSheetsTask()
    {
        JsonObject sheetsCollection = new JsonObject();

        SpreadsheetsResource.GetRequest getSpreadsheetRequest = service.Spreadsheets.Get(spreadsheetId);
        Spreadsheet spreadsheetResponse = await getSpreadsheetRequest.ExecuteAsync();

        if (spreadsheetResponse != null)
        {
            foreach (Sheet sheet in spreadsheetResponse.Sheets)
            {
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadsheetId, sheet.Properties.Title);
                request.MajorDimension = SpreadsheetsResource.ValuesResource.GetRequest.MajorDimensionEnum.ROWS;
                ValueRange valueResponse = await request.ExecuteAsync();

                JsonObject jsonSheet = new JsonObject();
                sheetsCollection.Add(sheet.Properties.Title, jsonSheet);

                if (valueResponse != null)
                {
                    IList<IList<System.Object>> values = valueResponse.Values;
                    if (values != null && values.Count > 1)
                    {
                        string[] keys = new string[values[0].Count];

                        for(int i = 0; i < values[0].Count; i++)
                        {
                            keys[i] = values[0][i].ToString();
                        }

                        for(int i = 1; i < values.Count; i++)
                        {
                            JsonObject jsonRow = new JsonObject();

                            for(int j = 1; j < values[i].Count; j++)
                            {
                                JsonValue jsonValue;

                                string valueString = values[i][j] as string;
                                if(valueString != null)
                                {
                                    bool tryBool;
                                    double tryDouble;

                                    if (bool.TryParse(valueString, out tryBool))
                                    {
                                        jsonValue = new JsonValue(tryBool);
                                    }
                                    else if (double.TryParse(valueString, out tryDouble))
                                    {
                                        jsonValue = new JsonValue(tryDouble);
                                    }
                                    else
                                    {
                                        jsonValue = new JsonValue(valueString);
                                    }
                                }
                                else
                                {
                                    jsonValue = new JsonValue();
                                } 
                                
                                jsonRow.Add(keys[j], jsonValue);
                            }

                            jsonSheet.Add(values[i][0].ToString(), jsonRow);
                        }
                    }
                }
            }
        }

        return sheetsCollection;
    }

    private void HandleApiException(GoogleApiException e)
    {
        string errorString = "";
        RequestError requestError = e.Error;
        if (requestError != null)
        {
            if (requestError.Errors != null)
            {
                errorString += "GoogleApiException - ";
                for (int i = 0; i < requestError.Errors.Count; i++)
                {
                    errorString += string.Format("Error {0} - Message: {1}" + (i + 1 < requestError.Errors.Count ? "\n" : ""), (i + 1).ToString(), requestError.Errors[i].Message);
                }
            }
        }

        if (string.IsNullOrWhiteSpace(errorString))
        {
            errorString = "GoogleApiException - Message: " + e.Message;
        }

        Debug.Log(errorString);
    }

    private void HandleException(Exception e)
    {
        Debug.Log("System Exception - Message: " + e.Message);
    }    
}
