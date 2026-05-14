# Security Review Guide

Security-focused code review checklist based on OWASP Top 10 and best practices.

## Authentication & Authorization

### Authentication
- [ ] Passwords hashed with strong algorithm (bcrypt, argon2)
- [ ] Password complexity requirements enforced
- [ ] Account lockout after failed attempts
- [ ] Secure password reset flow
- [ ] Multi-factor authentication for sensitive operations
- [ ] Session tokens are cryptographically random
- [ ] Session timeout implemented

### Authorization
- [ ] Authorization checks on every request
- [ ] Principle of least privilege applied
- [ ] Role-based access control (RBAC) properly implemented
- [ ] No privilege escalation paths
- [ ] Direct object reference checks (IDOR prevention)
- [ ] API endpoints protected appropriately

### JWT Security
```typescript
// ❌ Insecure JWT configuration
jwt.sign(payload, 'weak-secret');

// ✅ Secure JWT configuration
jwt.sign(payload, process.env.JWT_SECRET, {
  algorithm: 'RS256',
  expiresIn: '15m',
  issuer: 'your-app',
  audience: 'your-api'
});

// ❌ Not verifying JWT properly
const decoded = jwt.decode(token);  // No signature verification!

// ✅ Verify signature and claims
const decoded = jwt.verify(token, publicKey, {
  algorithms: ['RS256'],
  issuer: 'your-app',
  audience: 'your-api'
});
```

## Input Validation

### SQL Injection Prevention
```python
# ❌ Vulnerable to SQL injection
query = f"SELECT * FROM users WHERE id = {user_id}"

# ✅ Use parameterized queries
cursor.execute("SELECT * FROM users WHERE id = %s", (user_id,))

# ✅ Use ORM with proper escaping
User.objects.filter(id=user_id)
```

### XSS Prevention
```typescript
// ❌ Vulnerable to XSS
element.innerHTML = userInput;

// ✅ Use textContent for plain text
element.textContent = userInput;

// ✅ Use DOMPurify for HTML
element.innerHTML = DOMPurify.sanitize(userInput);

// ✅ React automatically escapes (but watch dangerouslySetInnerHTML)
return <div>{userInput}</div>;  // Safe
return <div dangerouslySetInnerHTML={{__html: userInput}} />;  // Dangerous!
```

### Command Injection Prevention
```python
# ❌ Vulnerable to command injection
os.system(f"convert {filename} output.png")

# ✅ Use subprocess with list arguments
subprocess.run(['convert', filename, 'output.png'], check=True)

# ✅ Validate and sanitize input
import shlex
safe_filename = shlex.quote(filename)
```

### Path Traversal Prevention
```typescript
// ❌ Vulnerable to path traversal
const filePath = `./uploads/${req.params.filename}`;

// ✅ Validate and sanitize path
const path = require('path');
const safeName = path.basename(req.params.filename);
const filePath = path.join('./uploads', safeName);

// Verify it's still within uploads directory
if (!filePath.startsWith(path.resolve('./uploads'))) {
  throw new Error('Invalid path');
}
```

## Data Protection

### Sensitive Data Handling
- [ ] No secrets in source code
- [ ] Secrets stored in environment variables or secret manager
- [ ] Sensitive data encrypted at rest
- [ ] Sensitive data encrypted in transit (HTTPS)
- [ ] PII handled according to regulations (GDPR, etc.)
- [ ] Sensitive data not logged
- [ ] Secure data deletion when required

### Configuration Security
```yaml
# ❌ Secrets in config files
database:
  password: "super-secret-password"

# ✅ Reference environment variables
database:
  password: ${DATABASE_PASSWORD}
```

### Error Messages
```typescript
// ❌ Leaking sensitive information
catch (error) {
  return res.status(500).json({
    error: error.stack,  // Exposes internal details
    query: sqlQuery      // Exposes database structure
  });
}

// ✅ Generic error messages
catch (error) {
  logger.error('Database error', { error, userId });  // Log internally
  return res.status(500).json({
    error: 'An unexpected error occurred'
  });
}
```

## API Security

### Rate Limiting
- [ ] Rate limiting on all public endpoints
- [ ] Stricter limits on authentication endpoints
- [ ] Per-user and per-IP limits
- [ ] Graceful handling when limits exceeded

### CORS Configuration
```typescript
// ❌ Overly permissive CORS
app.use(cors({ origin: '*' }));

// ✅ Restrictive CORS
app.use(cors({
  origin: ['https://your-app.com'],
  methods: ['GET', 'POST'],
  credentials: true
}));
```

### HTTP Headers
```typescript
// Security headers to set
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'"],
    }
  },
  hsts: { maxAge: 31536000, includeSubDomains: true },
  noSniff: true,
  xssFilter: true,
  frameguard: { action: 'deny' }
}));
```

## Cryptography

### Secure Practices
- [ ] Using well-established algorithms (AES-256, RSA-2048+)
- [ ] Not implementing custom cryptography
- [ ] Using cryptographically secure random number generation
- [ ] Proper key management and rotation
- [ ] Secure key storage (HSM, KMS)

### Common Mistakes
```typescript
// ❌ Weak random generation
const token = Math.random().toString(36);

// ✅ Cryptographically secure random
const crypto = require('crypto');
const token = crypto.randomBytes(32).toString('hex');

// ❌ MD5/SHA1 for passwords
const hash = crypto.createHash('md5').update(password).digest('hex');

// ✅ Use bcrypt or argon2
const bcrypt = require('bcrypt');
const hash = await bcrypt.hash(password, 12);
```

## Dependency Security

### Checklist
- [ ] Dependencies from trusted sources only
- [ ] No known vulnerabilities (npm audit, cargo audit)
- [ ] Dependencies kept up to date
- [ ] Lock files committed (package-lock.json, Cargo.lock)
- [ ] Minimal dependency usage
- [ ] License compliance verified

### Audit Commands
```bash
# Node.js
npm audit
npm audit fix

# Python
pip-audit
safety check

# Rust
cargo audit

# General
snyk test
```

## Logging & Monitoring

### Secure Logging
- [ ] No sensitive data in logs (passwords, tokens, PII)
- [ ] Logs protected from tampering
- [ ] Appropriate log retention
- [ ] Security events logged (login attempts, permission changes)
- [ ] Log injection prevented

```typescript
// ❌ Logging sensitive data
logger.info(`User login: ${email}, password: ${password}`);

// ✅ Safe logging
logger.info('User login attempt', { email, success: true });
```

## Security Review Severity Levels

| Severity | Description | Action |
|----------|-------------|--------|
| **Critical** | Immediate exploitation possible, data breach risk | Block merge, fix immediately |
| **High** | Significant vulnerability, requires specific conditions | Block merge, fix before release |
| **Medium** | Moderate risk, defense in depth concern | Should fix, can merge with tracking |
| **Low** | Minor issue, best practice violation | Nice to fix, non-blocking |
| **Info** | Suggestion for improvement | Optional enhancement |
