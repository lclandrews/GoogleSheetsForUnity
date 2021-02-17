// System
using System;
using System.Collections.Generic;
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
    public static GoogleSheetsManager instance
    {
        get
        {
            if (_instance == null)
            {
                GoogleSheetsManager[] managers = Resources.FindObjectsOfTypeAll<GoogleSheetsManager>();

                for (int i = 0; i < managers.Length; i++)
                {
                    if (managers[i].gameObject.scene.IsValid())
                    {
                        managers[i].Init();
                        break;
                    }
                }
            }

            if (_instance == null)
            {
                CreateManagerMonoBehavior();
            }

            return _instance;
        }
    }
    private static GoogleSheetsManager _instance;

    private static bool CreateManagerMonoBehavior()
    {
        if (_instance == null)
        {
            GameObject go = new GameObject("GoogleSheetsManager_Instance");
            GoogleSheetsManager manager = go.AddComponent<GoogleSheetsManager>();
            manager.Init();
            return true;
        }
        else
        {
            Debug.LogWarning("Cannot create manager mono behavior, a valid instance already exists.");
        }
        return false;
    }

    private static string applicationName = "Unity";

    [SerializeField]
    private string apiKey = "";
    [SerializeField]
    private string spreadsheetId = "";
    [SerializeField]
    private bool immediatelyRequestData = false;

    public bool initialized { get; private set; } = false;
    public bool retrievingData { get; private set; } = false;

    private SheetsRequestResponse onRequestResponse = null;

    private JsonObject sheets = null;

    private SheetsService service = null;    

    private void Awake()
    {
        if (_instance != null && _instance != this)
        {
            Debug.LogWarning("A Google Sheets manager MonoBehaviour was destroyed because there was more than one instance.");
            Destroy(this);
        }
        else
        {
            Init();
        }
    }    

    private void Update()
    {
        if(initialized && immediatelyRequestData)
        {
            immediatelyRequestData = false;
            GetSheets(null);
        }
    }

    private void Init()
    {
        if (!initialized)
        {
            _instance = this;
            DontDestroyOnLoad(gameObject);

            service = new SheetsService(new BaseClientService.Initializer()
            {
                ApiKey = apiKey,
                ApplicationName = applicationName,
            });

            initialized = true;
        }        
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
                                if(!string.IsNullOrWhiteSpace(valueString))
                                {
                                    if(!JsonValue.TryParse(valueString, out jsonValue))
                                    {
                                        jsonValue = new JsonValue(valueString);
                                    }
                                }
                                else
                                {
                                    jsonValue = JsonValue.Null;
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
