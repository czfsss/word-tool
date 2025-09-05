# Privacy Policy

## Overview

The word-tool plugin (hereinafter referred to as "this Plugin") is committed to protecting user privacy and data security. This privacy policy explains how we collect, use, store, and protect the information you provide when using this Plugin.

## Data Collection

### Types of Information We Collect

1. **Document Files**

   - PDF files (for PDF to Word conversion functionality)
   - Word documents (for annotation and chunk processing functionality)
   - Markdown files (for Markdown to Word conversion functionality)

2. **User Input Data**

   - Custom filenames
   - Annotation content and annotator information
   - Document processing parameter configurations

3. **System Log Information**
   - Processing operation records
   - Error log information
   - Performance monitoring data

### Purpose of Data Collection

- Provide document format conversion services
- Implement Word document annotation functionality
- Perform document chunk processing
- Improve plugin performance and user experience
- Troubleshoot and fix technical issues

## Data Processing

### Processing Methods

1. **Local Processing**

   - All document processing is performed in the local environment
   - Temporary files are used for intermediate processing and automatically deleted after completion
   - User documents are never uploaded to external servers

2. **In-Memory Processing**
   - Document content is temporarily stored in memory only during processing
   - Immediately cleared from memory after processing completion

### Technical Implementation

- Document processing based on python-docx and pdf2docx libraries
- Use dify_plugin framework to ensure secure plugin runtime environment
- Modular design to minimize data exposure scope

## Data Storage

### Storage Policy

1. **Temporary Storage**

   - Temporary files during processing are stored in system temporary directory
   - All temporary files are automatically deleted after processing completion
   - Original uploaded documents are never permanently saved

2. **Log Storage**

   - System logs contain only operation records and error information
   - Document content is never logged
   - Log files follow system default retention policies

3. **No Persistent Storage**
   - This Plugin does not permanently store user document content
   - No backups or caches of user documents are created

## Data Security

### Security Measures

1. **Access Control**

   - Plugin runs in a restricted environment
   - Only has necessary file system access permissions
   - Does not access other sensitive areas of user system

2. **Data Transmission**

   - All processing is performed locally with no network transmission
   - No user data is sent to third-party services
   - Plugin communication with external systems is limited to necessary framework interactions

3. **Memory Security**
   - Uses secure memory management mechanisms
   - Promptly releases memory resources that are no longer needed
   - Prevents memory leaks and data residue

### Security Best Practices

- Regularly update dependency libraries to fix security vulnerabilities
- Follow secure coding standards
- Implement principle of least privilege
- Use exception handling mechanisms to ensure data security

## Third-Party Dependencies

### Third-Party Libraries Used

1. **python-docx**: For Word document processing
2. **pdf2docx**: For PDF to Word conversion
3. **dify_plugin**: Provides plugin runtime framework
4. **Other standard Python libraries**: For file processing and data operations

### Third-Party Privacy Policies

All third-party libraries are open-source projects that do not collect or transmit user data. Users are recommended to review the official documentation of relevant libraries for detailed information.

## User Rights

### Your Rights

1. **Transparency Rights**

   - Understand how the Plugin processes your data
   - Obtain detailed information about data processing

2. **Control Rights**

   - Choose whether to use specific features
   - Control the scope of data provided to the Plugin

3. **Security Rights**
   - Expect data to be handled securely
   - Receive timely notification when security issues are discovered

### Data Minimization

We commit to:

- Collect only data necessary to complete functionality
- Not collect personal information unrelated to functionality
- Promptly delete temporary data that is no longer needed

## Data Retention

### Retention Policy

1. **Document Data**: Immediately deleted after processing completion, no retention
2. **Temporary Files**: Automatically cleaned, not exceeding single processing session
3. **Log Data**: Retained according to system default policies, contains no sensitive content
4. **Configuration Data**: Only retain non-sensitive parameters actively configured by users

## Children's Privacy

This Plugin is not specifically designed for children and does not intentionally collect personal information from children. If we discover that children's information has been collected, we will immediately delete the relevant data.

## Privacy Policy Updates

### Update Notifications

- We may update this privacy policy from time to time
- Significant changes will be communicated to users through appropriate means
- Users are recommended to regularly review the latest version of the privacy policy

### Version History

- **v1.0** (January 3, 2025): Initial version

## Contact Us

If you have any questions or concerns about this privacy policy, or need to exercise your privacy rights, please contact us through the following methods:

- **Project Repository**: Submit issues through GitHub Issues
- **Technical Support**: Review project documentation or contact maintainers

## Applicable Law

The interpretation and enforcement of this privacy policy are subject to relevant laws and regulations, including but not limited to:

- Personal information protection related laws and regulations
- Data security laws and regulations
- Network security laws and regulations

## Disclaimer

1. **Usage Risk**: Users should assess risks when using this Plugin to process sensitive documents
2. **Data Backup**: Users are recommended to backup important documents before use
3. **Environment Security**: Users should ensure the security of the runtime environment
4. **Compliant Use**: Users should ensure that use of this Plugin complies with relevant laws and regulations

---

**Last Updated**: January 3, 2025  
**Effective Date**: January 3, 2025  
**Version**: 1.0
