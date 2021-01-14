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

public struct AsyncOperation
{
    public static AsyncOperation empty { get { return new AsyncOperation(false, false, false); } }

    // Returns a new operation with an updated result, sets inProgress = true & broadcastResult = true
    public AsyncOperation ConstructUpdatedResult(bool resultSuccess) { return new AsyncOperation(resultSuccess, false, true); }
    // Returns a new operation with inProgress set to true
    public AsyncOperation ConstructInProgress() { return new AsyncOperation(this.resultSuccess, true, this.broadcastResult); }
    // Returns a new operation with broadcastResult set to false
    public AsyncOperation ConstructClearBroadcast() { return new AsyncOperation(this.resultSuccess, this.inProgress, false); }

    public bool resultSuccess { get; private set; }
    public bool inProgress { get; private set; }
    public bool broadcastResult { get; private set; }

    public bool ready { get { return !inProgress && !broadcastResult; } }

    public AsyncOperation(bool resultSuccess, bool inProgress, bool broadcastResult)
    {
        this.resultSuccess = resultSuccess;
        this.inProgress = inProgress;
        this.broadcastResult = broadcastResult;
    }
}

public class GoogleSheetsManager : MonoBehaviour
{
    private static string applicationName = "Unity";

    public string apiKey = "AIzaSyAbY5-EjROctQUECA8RWai9q_c8Uzlusvc";
    public string spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
    public bool immediatelyRequestData = false;

    public AsyncOperation initialized { get; private set; } = AsyncOperation.empty;
    public AsyncOperation retrievingData { get; private set; } = AsyncOperation.empty;

    private ServiceInitializedResponse onInitialized = null;
    private SheetsRequestResponse onRequestResponse = null;

    private JsonObject sheets = null;

    private SheetsService service = null;

    private string fullCredentialPath = "";
    private string fullTokenPath = "";

    private void Awake()
    {
        InitializeServiceApiKey(null);
    }

    private void Update()
    {
        if(initialized.broadcastResult)
        {
            initialized = initialized.ConstructClearBroadcast();
            onInitialized?.Invoke(initialized.resultSuccess);
            onInitialized = null;
        }

        if(retrievingData.broadcastResult)
        {
            retrievingData = retrievingData.ConstructClearBroadcast();
            onRequestResponse?.Invoke(retrievingData.resultSuccess, sheets);
            onRequestResponse = null;
        }

        if(initialized.resultSuccess && immediatelyRequestData)
        {
            immediatelyRequestData = false;
            GetSheets(null);
        }
    }

    private bool InitializeServiceApiKey(ServiceInitializedResponse response)
    {
        if (!initialized.resultSuccess && initialized.ready)
        {
            // Excessive usage of setting the operation struct
            // legacy from initializing async with oauth            
            initialized = initialized.ConstructInProgress();
            onInitialized += response;

            service = new SheetsService(new BaseClientService.Initializer()
            {
                ApiKey = apiKey,
                ApplicationName = applicationName,
            });

            initialized = initialized.ConstructUpdatedResult(true);
            return true;
        }
        else
        {
            onInitialized += response;
        }        
        return false;
    }

    public bool GetSheets(SheetsRequestResponse response)
    {
        if(initialized.resultSuccess && retrievingData.ready)
        {
            retrievingData = retrievingData.ConstructInProgress();
            onRequestResponse += response;
            ThreadPool.QueueUserWorkItem(GetSheetsInternal);
            return true;
        }
        else
        {
            onRequestResponse += response;
        }
        return false;
    }

    private void GetSheetsInternal(System.Object stateInfo)
    {       
        try
        {
            // Need to identify whether or not an exception / what exception is thrown for http errors
            Task sheetsTask = GetSheetsTask();
            sheetsTask.Wait();
        }
        catch (AggregateException ae)
        {
            ae.Handle((Exception e) =>
            {
                if(e is GoogleApiException)
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
        catch(Exception e)
        {
            HandleException(e);
        }
    }

    private async Task GetSheetsTask()
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
                                if (values[i][j] is bool)
                                {
                                    jsonValue = new JsonValue((bool)values[i][j]);
                                }
                                else if (values[i][j] is double)
                                {
                                    jsonValue = new JsonValue((double)values[i][j]);
                                }
                                else if (values[i][j] is string)
                                {
                                    jsonValue = new JsonValue((string)values[i][j]);
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

        sheets = sheetsCollection;
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
