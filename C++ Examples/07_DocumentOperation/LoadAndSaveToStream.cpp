#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"LoadAndSaveToStream.rtf";

	ifstream inputf(inputFile.c_str(), ios::in | ios::binary);
	// Open the stream. Read only access is enough to load a document.
	Stream* stream = new Stream(inputf);

	// Load the entire document into memory.
	Document* doc = new Document(stream);

	// You can close the stream now, it is no longer needed because the document is in memory.
	stream->Close();
	// Do something with the document

	// Convert the document to a different format and save to stream.
	Stream* newStream = new Stream();
	doc->SaveToStream(newStream, FileFormat::Rtf);

	newStream->Save(outputFile.c_str());

	doc->Close();
	delete doc;
}
