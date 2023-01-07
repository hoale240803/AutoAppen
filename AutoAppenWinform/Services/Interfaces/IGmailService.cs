namespace AutoAppenWinform.Services.Interfaces
{
    public interface IGmailService
    {
        void SayHello();

        void GetUnreadMessages();

        void Output(string output);
        Task<string> GetAccessToken();

        Task<string> GetGmailVerificationCode(string access_token);

        string CreateGmailOAth2AuthorizationRequest(string redirectURI, string state, string code_challenge, string code_challenge_method);

        void LoginGmailAccount();
    }
}