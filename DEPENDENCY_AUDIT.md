# Dependency Audit Guide

This guide provides a comprehensive approach to auditing dependencies for outdated packages, security vulnerabilities, and unnecessary bloat across different technology stacks.

## Table of Contents
- [Node.js / JavaScript](#nodejs--javascript)
- [Python](#python)
- [Ruby](#ruby)
- [Java](#java)
- [Go](#go)
- [General Best Practices](#general-best-practices)
- [Audit Checklist](#audit-checklist)

---

## Node.js / JavaScript

### Check for Outdated Packages

```bash
# List outdated packages
npm outdated

# Check for major updates
npx npm-check-updates

# Interactive update tool
npx npm-check -u
```

### Security Vulnerabilities

```bash
# Run npm audit
npm audit

# Fix automatically (use with caution)
npm audit fix

# Fix breaking changes manually
npm audit fix --force

# Alternative: Use Snyk
npx snyk test
npx snyk wizard
```

### Analyze Bundle Size & Bloat

```bash
# Analyze what's taking up space
npx depcheck                    # Find unused dependencies
npx cost-of-modules             # Show size of node_modules
npx webpack-bundle-analyzer     # For webpack projects
npx source-map-explorer         # Analyze bundle composition

# Check duplicate dependencies
npm dedupe
npm ls
```

### Recommendations
- **Keep dependencies minimal**: Only install what you actually use
- **Use exact versions** in package.json for critical dependencies
- **Prefer smaller alternatives**:
  - Use `date-fns` instead of `moment` (smaller)
  - Use `axios` instead of `request` (maintained)
  - Consider native solutions before adding libraries
- **Regular maintenance**: Update dependencies at least monthly
- **Lock files**: Always commit package-lock.json or yarn.lock

---

## Python

### Check for Outdated Packages

```bash
# List outdated packages
pip list --outdated

# Using pip-review
pip install pip-review
pip-review --local --interactive

# Check specific package
pip show <package-name>
```

### Security Vulnerabilities

```bash
# Using safety
pip install safety
safety check
safety check --json

# Using pip-audit
pip install pip-audit
pip-audit

# Snyk for Python
pip install snyk
snyk test --file=requirements.txt
```

### Analyze Dependencies

```bash
# Show dependency tree
pip install pipdeptree
pipdeptree

# Find unused dependencies
pip install pip-autoremove
pip-autoremove <package-name> --list

# Check for conflicting dependencies
pip check
```

### Recommendations
- **Use virtual environments**: Always isolate project dependencies
- **Pin versions**: Use `requirements.txt` with exact versions
- **Use requirements files**:
  - `requirements.txt` - production dependencies
  - `requirements-dev.txt` - development dependencies
- **Regular updates**: Review and update quarterly
- **Consider alternatives**:
  - Use `httpx` instead of `requests` (async support)
  - Use built-in `json` instead of external libraries when possible

---

## Ruby

### Check for Outdated Packages

```bash
# List outdated gems
bundle outdated

# Check specific gem
gem list <gem-name>
```

### Security Vulnerabilities

```bash
# Using bundler-audit
gem install bundler-audit
bundle audit check
bundle audit update

# Alternative: brakeman (for Rails)
gem install brakeman
brakeman
```

### Analyze Dependencies

```bash
# Show dependency tree
bundle viz

# List all dependencies
bundle list
```

### Recommendations
- **Gemfile.lock**: Always commit to version control
- **Semantic versioning**: Use `~>` operator carefully
- **Regular updates**: Update gems monthly
- **Security first**: Subscribe to Ruby security mailing lists

---

## Java

### Check for Outdated Packages (Maven)

```bash
# Maven versions plugin
mvn versions:display-dependency-updates
mvn versions:display-plugin-updates

# Maven dependency tree
mvn dependency:tree
```

### Check for Outdated Packages (Gradle)

```bash
# Gradle versions plugin
./gradlew dependencyUpdates

# Dependency tree
./gradlew dependencies
```

### Security Vulnerabilities

```bash
# OWASP Dependency Check (Maven)
mvn org.owasp:dependency-check-maven:check

# OWASP Dependency Check (Gradle)
./gradlew dependencyCheckAnalyze

# Snyk
snyk test --file=pom.xml
snyk test --file=build.gradle
```

### Recommendations
- **Use dependency management**: Centralize versions in parent POM
- **Exclude transitive dependencies**: Remove unused transitive deps
- **Regular updates**: Quarterly dependency reviews
- **Security scanning**: Integrate OWASP checks into CI/CD

---

## Go

### Check for Outdated Packages

```bash
# List available updates
go list -u -m all

# Using go-mod-outdated
go install github.com/psampaz/go-mod-outdated@latest
go list -u -m -json all | go-mod-outdated

# Update all dependencies
go get -u ./...
```

### Security Vulnerabilities

```bash
# Go vulnerability check
go install golang.org/x/vuln/cmd/govulncheck@latest
govulncheck ./...

# Nancy (Sonatype)
go list -json -m all | nancy sleuth
```

### Analyze Dependencies

```bash
# Dependency graph
go mod graph

# Why is this dependency here?
go mod why <package>

# Tidy up
go mod tidy

# Verify dependencies
go mod verify
```

### Recommendations
- **Minimal dependencies**: Go encourages standard library usage
- **go.sum**: Always commit to version control
- **Vendor if needed**: Consider vendoring for critical projects
- **Regular maintenance**: Run `go mod tidy` regularly
- **Direct dependencies only**: Avoid unnecessary indirect deps

---

## General Best Practices

### 1. Automated Dependency Management

**Tools to Consider:**
- **Dependabot** (GitHub): Automated dependency updates
- **Renovate**: Highly configurable dependency updates
- **Snyk**: Security-focused dependency management
- **WhiteSource**: Enterprise dependency management

### 2. Security Best Practices

- ✅ Enable automated security alerts (GitHub/GitLab)
- ✅ Review security advisories regularly
- ✅ Subscribe to security mailing lists for your stack
- ✅ Use Software Composition Analysis (SCA) tools
- ✅ Implement dependency review in code review process
- ✅ Monitor CVE databases for your dependencies

### 3. Update Strategy

**Semantic Versioning (SemVer):**
- **Patch** (1.0.x): Bug fixes - safe to update
- **Minor** (1.x.0): New features - usually safe
- **Major** (x.0.0): Breaking changes - requires testing

**Update Cadence:**
- **Security patches**: Immediately
- **Patch updates**: Weekly/bi-weekly
- **Minor updates**: Monthly
- **Major updates**: Quarterly (with thorough testing)

### 4. Bloat Prevention

**Questions to Ask:**
- Do I really need this dependency?
- Can I implement this functionality myself in <50 lines?
- Is there a lighter alternative?
- Am I using most of the features, or just one function?
- Is this dependency actively maintained?

**Red Flags:**
- ⚠️ Package hasn't been updated in >2 years
- ⚠️ Many open security issues
- ⚠️ Large size for simple functionality
- ⚠️ Lots of transitive dependencies
- ⚠️ Deprecated or archived repository

---

## Audit Checklist

### Initial Setup
- [ ] Lock files are committed to version control
- [ ] CI/CD includes dependency security scanning
- [ ] Automated dependency update tool configured (Dependabot/Renovate)
- [ ] Security alerts enabled on repository

### Monthly Tasks
- [ ] Review and update patch versions
- [ ] Check for security vulnerabilities
- [ ] Review automated dependency update PRs
- [ ] Remove unused dependencies

### Quarterly Tasks
- [ ] Review all outdated dependencies
- [ ] Evaluate and update minor versions
- [ ] Audit dependency tree for bloat
- [ ] Review and update major versions (with testing)
- [ ] Document breaking changes

### Audit Questions
1. **Is it outdated?**
   - When was the last update?
   - Is there a newer major version?
   - Is the package deprecated?

2. **Is it secure?**
   - Any known CVEs?
   - Is it actively maintained?
   - Does it have security best practices?

3. **Is it necessary?**
   - Are we actually using it?
   - Could we use a native/built-in alternative?
   - Is there a lighter alternative?
   - Are we only using a tiny part of it?

4. **Is it quality?**
   - Good documentation?
   - Active community?
   - Good test coverage?
   - Stable API?

### Documentation
- [ ] Document why each major dependency was chosen
- [ ] Keep an architecture decision record (ADR) for dependencies
- [ ] Document known vulnerabilities and mitigation strategies
- [ ] Maintain a list of approved/banned dependencies

---

## Quick Commands Reference

### Node.js
```bash
npm outdated && npm audit && npx depcheck
```

### Python
```bash
pip list --outdated && pip-audit && pipdeptree
```

### Ruby
```bash
bundle outdated && bundle audit check
```

### Java (Maven)
```bash
mvn versions:display-dependency-updates && mvn org.owasp:dependency-check-maven:check
```

### Java (Gradle)
```bash
./gradlew dependencyUpdates && ./gradlew dependencyCheckAnalyze
```

### Go
```bash
go list -u -m all && govulncheck ./... && go mod tidy
```

---

## Additional Resources

- [OWASP Dependency Check](https://owasp.org/www-project-dependency-check/)
- [Snyk Vulnerability Database](https://snyk.io/vuln)
- [GitHub Advisory Database](https://github.com/advisories)
- [National Vulnerability Database](https://nvd.nist.gov/)
- [npm security best practices](https://docs.npmjs.com/packages-and-modules/securing-your-code)
- [Python security guide](https://python.readthedocs.io/en/stable/library/security_warnings.html)
