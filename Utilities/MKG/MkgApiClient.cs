using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq; // FIXED: Added missing using directive

public class MkgApiClient : IDisposable
{
    private readonly HttpClient _client;
    private string _sessionId;
    private readonly string _baseUrl;
    private readonly string _authUrl;
    private readonly string _restUrl;
    private readonly string _username;
    private readonly string _password;
    private readonly string _apiKey;

    public MkgApiClient()
    {
        // Create HttpClient with proper setup
        var handler = new HttpClientHandler() { UseCookies = false }; // Don't use auto cookie handling
        _client = new HttpClient(handler);

        // Read from App.config
        _baseUrl = ConfigurationManager.AppSettings["MkgApi:Urls:Base"];
        _authUrl = ConfigurationManager.AppSettings["MkgApi:Urls:Auth"];
        _restUrl = ConfigurationManager.AppSettings["MkgApi:Urls:Rest"];
        _username = ConfigurationManager.AppSettings["MkgApi:Username"];
        _password = ConfigurationManager.AppSettings["MkgApi:Password"];
        _apiKey = ConfigurationManager.AppSettings["MkgApi:ApiKey"];

        // Set default headers
        _client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    }

    /// <summary>
    /// Voert een login uit bij MKG en bewaart de sessie-cookie
    /// </summary>
    public async Task<bool> LoginAsync()
    {
        try
        {
            // Create form data for login
            var formData = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("j_username", _username),
                new KeyValuePair<string, string>("j_password", _password)
            });

            // Build full auth URL
            string fullAuthUrl = $"{_baseUrl}{_authUrl}";
            Console.WriteLine($">>> AUTH URL: {fullAuthUrl}");
            Console.WriteLine($">>> Username: {_username}");
            Console.WriteLine($">>> Password: {(_password?.Length > 0 ? "***SET***" : "***EMPTY***")}");

            // Perform login POST
            var response = await _client.PostAsync(fullAuthUrl, formData);
            string responseBody = await response.Content.ReadAsStringAsync();

            Console.WriteLine($">>> Login response status: {response.StatusCode}");
            Console.WriteLine($">>> Login response body: {responseBody}");

            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine(">>> Login failed - HTTP error");
                Console.WriteLine($">>> Response headers:");
                foreach (var header in response.Headers)
                {
                    Console.WriteLine($"    {header.Key}: {string.Join(", ", header.Value)}");
                }
                return false;
            }

            // FIXED: Extract session ID from Set-Cookie headers
            if (response.Headers.TryGetValues("Set-Cookie", out var cookies))
            {
                // Convert to list to fix the Count() issue
                var cookieList = cookies.ToList();
                Console.WriteLine($">>> Found {cookieList.Count} cookies in response");

                foreach (var cookie in cookieList)
                {
                    Console.WriteLine($">>> Cookie: {cookie}");
                    if (cookie.StartsWith("JSESSIONID"))
                    {
                        var parts = cookie.Split(';')[0].Split('=');
                        if (parts.Length == 2)
                        {
                            _sessionId = parts[1];
                            Console.WriteLine($">>> Login successful, Session ID: {_sessionId}");
                            return true;
                        }
                    }
                }
            }

            Console.WriteLine(">>> Login response received but no JSESSIONID found");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($">>> Login exception: {ex.Message}");
            Console.WriteLine($">>> Stack trace: {ex.StackTrace}");
            return false;
        }
    }

    /// <summary>
    /// Voert een GET-request uit op het MKG REST endpoint
    /// </summary>
    public async Task<string> GetAsync(string endpoint)
    {
        if (string.IsNullOrEmpty(_sessionId))
        {
            var success = await LoginAsync();
            if (!success)
                throw new HttpRequestException("MKG login failed.");
        }

        // Build full URL
        string fullUrl = $"{_baseUrl}{_restUrl}/{endpoint}";

        var request = new HttpRequestMessage(HttpMethod.Get, fullUrl);
        request.Headers.Add("X-CustomerID", _apiKey);
        request.Headers.Add("Accept", "application/json");
        request.Headers.Add("Cookie", $"JSESSIONID={_sessionId}");

        var response = await _client.SendAsync(request);
        response.EnsureSuccessStatusCode();

        return await response.Content.ReadAsStringAsync();
    }

    public async Task<string> PostAsync(string endpoint, HttpContent content)
    {
        if (string.IsNullOrEmpty(_sessionId))
        {
            var success = await LoginAsync();
            if (!success)
                throw new HttpRequestException("MKG login failed.");
        }

        // Build full URL - FIXED: Remove extra slashes and ensure proper endpoint format
        string cleanEndpoint = endpoint.TrimStart('/');
        string fullUrl = $"{_baseUrl}{_restUrl}/{cleanEndpoint}";

        var request = new HttpRequestMessage(HttpMethod.Post, fullUrl);
        request.Headers.Add("X-CustomerID", _apiKey);
        request.Headers.Add("Accept", "application/json");
        request.Headers.Add("Cookie", $"JSESSIONID={_sessionId}");
        request.Content = content;

        Console.WriteLine($">>> POST URL: {fullUrl}");
        Console.WriteLine($">>> Request Headers:");
        Console.WriteLine($"    X-CustomerID: {_apiKey}");
        Console.WriteLine($"    Accept: application/json");
        Console.WriteLine($"    Cookie: JSESSIONID={_sessionId}");

        // Log the request body for debugging
        if (content != null)
        {
            string requestBody = await content.ReadAsStringAsync();
            Console.WriteLine($">>> Request Body:\n{requestBody}");
        }

        var response = await _client.SendAsync(request);

        // READ THE RESPONSE BODY BEFORE CHECKING SUCCESS
        string responseBody = await response.Content.ReadAsStringAsync();

        Console.WriteLine($">>> Response Status: {response.StatusCode}");
        Console.WriteLine($">>> Response Headers:");
        foreach (var header in response.Headers)
        {
            Console.WriteLine($"    {header.Key}: {string.Join(", ", header.Value)}");
        }
        Console.WriteLine($">>> Response Body:\n{responseBody}");

        // FIXED: Better error handling for HTML responses
        if (!response.IsSuccessStatusCode)
        {
            if (responseBody.Contains("<!doctype html>") || responseBody.Contains("<html"))
            {
                // Extract error message from HTML if possible
                var errorMessage = ExtractErrorFromHtml(responseBody);
                throw new HttpRequestException($"HTTP {(int)response.StatusCode} {response.StatusCode}: {errorMessage}");
            }
            else
            {
                throw new HttpRequestException($"HTTP {(int)response.StatusCode} {response.StatusCode}: {responseBody}");
            }
        }

        return responseBody;
    }

    public async Task<string> PutAsync(string endpoint, HttpContent content)
    {
        if (string.IsNullOrEmpty(_sessionId))
        {
            var success = await LoginAsync();
            if (!success)
                throw new HttpRequestException("MKG login failed.");
        }

        // Build full URL
        string fullUrl = $"{_baseUrl}{_restUrl}/{endpoint}";

        var request = new HttpRequestMessage(HttpMethod.Put, fullUrl);
        request.Headers.Add("X-CustomerID", _apiKey);
        request.Headers.Add("Accept", "application/json");
        request.Headers.Add("Cookie", $"JSESSIONID={_sessionId}");
        request.Content = content;

        var response = await _client.SendAsync(request);
        response.EnsureSuccessStatusCode();

        return await response.Content.ReadAsStringAsync();
    }

    public async Task<string> DeleteAsync(string endpoint)
    {
        if (string.IsNullOrEmpty(_sessionId))
        {
            var success = await LoginAsync();
            if (!success)
                throw new HttpRequestException("MKG login failed.");
        }

        // Build full URL
        string fullUrl = $"{_baseUrl}{_restUrl}/{endpoint}";

        var request = new HttpRequestMessage(HttpMethod.Delete, fullUrl);
        request.Headers.Add("X-CustomerID", _apiKey);
        request.Headers.Add("Accept", "application/json");
        request.Headers.Add("Cookie", $"JSESSIONID={_sessionId}");

        var response = await _client.SendAsync(request);
        response.EnsureSuccessStatusCode();

        return await response.Content.ReadAsStringAsync();
    }

    // Helper method to extract error messages from HTML responses
    private string ExtractErrorFromHtml(string htmlResponse)
    {
        try
        {
            // Try to extract meaningful error message from HTML
            if (htmlResponse.Contains("Bad Request"))
            {
                return "Bad Request - Check API payload format and required fields";
            }
            else if (htmlResponse.Contains("Unauthorized"))
            {
                return "Unauthorized - Check API credentials and session";
            }
            else if (htmlResponse.Contains("Not Found"))
            {
                return "Not Found - Check API endpoint and resource existence";
            }
            else
            {
                return "Server returned HTML error page - check logs for details";
            }
        }
        catch
        {
            return "Unknown HTML error response";
        }
    }

    public void Dispose()
    {
        _client?.Dispose();
    }
}