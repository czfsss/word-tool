## Word Document Processing Toolkit

A comprehensive Word document processing plugin collection designed specifically for the Dify platform, providing PDF document conversion, intelligent chunking, and annotation features. Supports Chinese document processing with localized processing strategies to ensure data security.

## 🚀 Core Features

### 1. PDF to Word Converter (pdf_to_word)

- **Function Description**: Convert PDF documents to editable Word documents
- **Technical Implementation**: Based on `pdf2docx` library, supports complex layouts and format preservation
- **Supported Formats**: PDF → DOCX
- **Custom Features**: Support for custom output file names

### 2. Word Intelligent Chunker (word_chunk) ⭐

- **Function Description**: Intelligently analyze the structure of Word documents and perform reasonable document chunking based on potential headings, trying to ensure that sentence segmentation during document chunking does not cause incomplete semantic recognition by models. Note: The Word format should be as standardized as possible, and it works better for contract and legal documents.
- **Application Scenarios**: Document analysis, content extraction, knowledge management
- **Chunking Strategy**: Intelligent chunking algorithm based on paragraph semantics and document structure
- **Quantity Control**: Support for custom chunk numbers (maximum 30), with the 30-chunk limit set because iteration nodes can have at most 30 iteration objects.

### 3. Word Document Annotator (word_comment)

- **Function Description**: Add genuine native comments to Word documents
- **Technical Features**: Uses `python-docx 1.2.0` native comment API
- **Comment Features**: Precise text positioning, format preservation, support for custom commenters
- **Input Method**: JSON format comment data with flexible configuration

## 📊 Word Intelligent Chunking Algorithm Details

### Algorithm Overview

The Word intelligent chunker employs multi-level document analysis strategies, combining semantic understanding and structural recognition to achieve intelligent document segmentation. This algorithm is specifically optimized for contract, legal, and policy documents, intelligently identifying document structures and performing reasonable chunking to ensure that document segmentation does not cause incomplete semantic recognition by models.

### Core Algorithm Flow

1. **Document Element Extraction and Preprocessing**: Extract paragraph and table content in the original document order, preserving document structure information.
2. **Document Type Recognition**: Identify document types (general, contract, policy) and adopt specific processing strategies for different types.
3. **Intelligent Title Recognition**: Recognize titles through multiple dimensions including style names, regex pattern matching, format feature analysis, and document type-specific patterns.
4. **Semantic Chunking Processing**: Adopt different chunking strategies based on different element types such as titles, tables, long paragraphs, short paragraphs, and special paragraphs.
5. **Chunk Quantity Control**: When the number of chunks exceeds the limit, use a mathematical grouping algorithm for intelligent merging.

### Algorithm Features

1. **Document Type Optimization**: Optimized for specific document types like contracts and policies, recognizing document-specific structures.
2. **Multi-dimensional Title Detection**: Combines style names, regex patterns, format features, and document type-specific patterns for title recognition.
3. **Intelligent Semantic Chunking**: Adopts different chunking strategies based on document element types to maintain semantic integrity.
4. **Special Table Processing**: Ensures tables are in the same chunk as related titles, maintaining contextual relevance.
5. **Precise Quantity Control**: Uses mathematical grouping algorithms to ensure precise chunk quantity control, avoiding over-chunking or under-chunking.
6. **Hierarchical Structure Recognition**: Capable of recognizing document hierarchical structures and maintaining logical integrity.

## 📖 Usage Guide

### PDF to Word Conversion Tool

#### Function Description

This tool converts PDF documents to Word format while preserving the original layout and formatting as much as possible. It utilizes the pdf2docx library for efficient conversion.

#### Input Parameters

- **file_path** (string, required): Path to the PDF file to be converted
- **output_path** (string, optional): Path for the output Word file. If not provided, a default path will be generated

#### Output Parameters

- **status** (string): Operation status ("success" or "error")
- **message** (string): Detailed message about the operation result
- **output_file** (string): Path to the generated Word file

### Word Intelligent Chunking Tool

#### Function Description

This tool intelligently segments Word documents into meaningful chunks, optimized for contract, legal, and policy documents. It identifies document structure, recognizes titles, and performs semantic chunking to maintain context integrity.

#### Input Parameters

- **file_path** (string, required): Path to the Word file to be chunked
- **max_chunk** (integer, optional, default: 30): Maximum number of chunks to generate
- **min_length** (integer, optional, default: 1000): Minimum character length for a chunk to be considered independent
- **chunk_overlap** (integer, optional, default: 200): Character overlap between adjacent chunks

#### Output Parameters

- **status** (string): Operation status ("success" or "error")
- **chunks** (array): List of generated chunks
  - **index** (integer): Chunk index
  - **content** (string): Text content of the chunk
  - **word_count** (integer): Word count in the chunk
- **total_chunks** (integer): Total number of chunks generated

### Word Document Commenting Tool

#### Function Description

This tool adds native comments to Word documents. It supports adding comments to specific text segments in paragraphs and tables, with intelligent text matching capabilities to locate the target text even when exact matches are not found.

#### Input Parameters

- **file_path** (string, required): Path to the Word file to be commented
- **comments** (array or object, required): List of comments to add, supports two formats:
  - **Format One (Object Array)**:
    ```json
    [
      {
        "text": "Comment content 1",
        "target_text": "Target text 1"
      },
      {
        "text": "Comment content 2",
        "target_text": "Target text 2"
      }
    ]
    ```
  - **Format Two (Object Array with Multiple Key-Value Pairs)**:
    ```json
    [
      {
        "Target text 1": "Comment content 1",
        "Target text 2": "Comment content 2"
      },
      {
        "Target text 1": "Comment content 1",
        "Target text 2": "Comment content 2"
      }
    ]
    ```
  - **text** (string, required): Comment text content (Format One)
  - **target_text** (string, required): Target text in the document to comment on (Format One)
  - **page** (integer, optional): Page number where the comment should be added (for table comments)

#### Output Parameters

- **status** (string): Operation status ("success" or "error")
- **message** (string): Detailed message about the operation result
- **output_file** (string): Path to the generated Word file with comments
- **invalid_comments** (array, optional): List of comments that could not be added
  - **text** (string): Comment text that failed
  - **target_text** (string): Target text that couldn't be found
  - **reason** (string): Reason for failure

## 🔧 Technical Architecture

### Core Technology Stack

- **Document Processing**: python-docx, pdf2docx
- **Plugin Framework**: dify_plugin
- **Logging System**: Unified logging and exception handling
- **File Processing**: Temporary file management and automatic cleanup

### Security Features

- **Local Processing**: All operations executed in local environment
- **Temporary Files**: Automatic cleanup after processing completion
- **Memory Safety**: Timely release of memory resources
- **Permission Control**: Principle of least privilege

### Performance Optimization

- **Memory Management**: Large document fragment processing
- **Algorithm Optimization**: Intelligent chunking algorithm with O(n) complexity
- **Caching Strategy**: Temporary result caching
- **Concurrent Processing**: Support for multi-document parallel processing

## 📊 Use Cases

### 1. Knowledge Management Systems

- Document import and preprocessing
- Content chunking and index establishment
- Comment and review management

### 2. Content Analysis Platforms

- Large document intelligent segmentation
- Structured data extraction
- Semantic analysis preprocessing

### 3. Collaborative Office Scenarios

- Document format conversion
- Comment and review workflows
- Version control and tracking

### 4. Automated Workflows

- Batch document processing
- Format standardization
- Content quality inspection

## 🐛 Troubleshooting

### Common Issues

**1. PDF Conversion Failure**

```
Cause: PDF file corrupted or encrypted
Solution: Check PDF file integrity, remove password protection
```

**2. Unsatisfactory Chunking Results**

```
Cause: Irregular document structure. This plugin's chunking is intended for standard format contract or policy documents
Solution: Check title formatting
```

**3. Comment Addition Failure**

```
Cause: Target text not found
Solution: Ensure target text exists in document and format matches
```

## 🚀 Changelog

### v0.0.1 (2025-01-03)

- ✨ New Feature: PDF to Word converter
- ✨ New Feature: Word intelligent chunker
- ✨ New Feature: Word document annotator
- 🔧 Optimization: Unified file processing logic
- 🔧 Optimization: Improved error handling mechanism
- 📝 Documentation: Enhanced usage guide and API documentation

---

**Developer**: czfsss  
**Version**: 0.0.1  
**Last Updated**: September 3, 2025
