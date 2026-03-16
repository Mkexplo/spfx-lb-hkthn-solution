# Azure OpenAI API Key Security Guide

## Overview
This document outlines best practices for securely managing Azure OpenAI API credentials in your SPFx application.

## ⚠️ NEVER DO THIS
- ❌ Hardcode API keys in source code
- ❌ Commit API keys to git repository
- ❌ Store keys in client-side code
- ❌ Log sensitive credentials
- ❌ Pass keys through unsecured HTTP requests

## ✅ RECOMMENDED APPROACHES

### Option 1: Environment Variables (Development)
**Best for:** Local development and testing

1. Create a `.env` file in your project root (don't commit this!)
   ```
   AZURE_OPENAI_ENDPOINT=https://your-resource.cognitiveservices.azure.com/
   AZURE_OPENAI_API_KEY=your-secret-key-here
   AZURE_OPENAI_API_VERSION=2025-01-01-preview
   AZURE_OPENAI_DEPLOYMENT=o4-mini
   ```

2. Install dotenv package:
   ```bash
   npm install dotenv
   ```

3. Load in your code:
   ```typescript
   import dotenv from 'dotenv';
   dotenv.config();
   
   const apiKey = process.env.AZURE_OPENAI_API_KEY;
   ```

4. Add `.env` to `.gitignore`:
   ```
   .env
   .env.local
   *.env
   ```

### Option 2: Azure Key Vault (Production - RECOMMENDED)
**Best for:** Production deployments

1. Install required packages:
   ```bash
   npm install @azure/identity @azure/keyvault-secrets
   ```

2. Use DefaultAzureCredential for authentication:
   ```typescript
   import { DefaultAzureCredential } from "@azure/identity";
   import { SecretClient } from "@azure/keyvault-secrets";
   
   const credential = new DefaultAzureCredential();
   const client = new SecretClient(
     "https://your-keyvault-name.vault.azure.net/",
     credential
   );
   
   const secret = await client.getSecret("AZURE_OPENAI_API_KEY");
   const apiKey = secret.value;
   ```

3. Configure Azure AD authentication for your SPFx app

### Option 3: Backend API Proxy (Recommended for SPFx)
**Best for:** SPFx web parts

Instead of calling Azure OpenAI directly from the client, call your own backend API:

```typescript
// In your SPFx component - only calls your backend
const response = await fetch('/api/ai/response', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${token}`
  },
  body: JSON.stringify({ message: userMessage })
});
```

Your backend handles the Azure OpenAI API call securely with stored credentials.

## SharePoint Online Considerations

For SPFx solutions deployed to SharePoint Online:

1. **Use Azure AD Authentication:**
   - Configure your Azure AD app registration
   - Use `aadTokenProviderFactory` (already in your code)
   - Token-based authentication is more secure

2. **Store secrets in Key Vault:**
   - Access Key Vault through Managed Identity
   - Never store credentials in SharePoint lists

3. **API Permissions:**
   - Limit API scope to only what's needed
   - Use role-based access control (RBAC)

## Implementation in spservice.ts

The current implementation uses environment variables with validation:

```typescript
const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
const apiKey = process.env.AZURE_OPENAI_API_KEY;

if (!endpoint || !apiKey) {
  console.error('Credentials not configured');
  return 'Configuration error';
}
```

To switch to Key Vault, replace the environment variable section with Key Vault retrieval.

## Security Checklist

- [ ] API keys are NOT in source code
- [ ] `.env` file is in `.gitignore`
- [ ] `.env.example` is committed as a template
- [ ] Production uses Azure Key Vault or backend proxy
- [ ] API calls are authenticated via bearer token
- [ ] Sensitive data is not logged
- [ ] Rate limiting is implemented
- [ ] API usage is monitored
- [ ] Credentials are rotated regularly

## Monitoring & Alerts

1. Set up Azure Monitor alerts for:
   - Failed authentication attempts
   - Unusual API usage patterns
   - Rate limit violations

2. Enable audit logging for:
   - All AI API calls
   - Configuration changes
   - Access to secrets

## Additional Resources

- [Azure Key Vault Best Practices](https://learn.microsoft.com/en-us/azure/key-vault/general/best-practices)
- [Azure OpenAI Security](https://learn.microsoft.com/en-us/azure/ai-services/openai/concepts/security)
- [SPFx Security Best Practices](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/security-considerations)
- [OWASP Secrets Management](https://cheatsheetseries.owasp.org/cheatsheets/Secrets_Management_Cheat_Sheet.html)
