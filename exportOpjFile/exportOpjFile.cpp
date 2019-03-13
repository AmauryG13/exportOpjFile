// exportOpjFile.cpp : Ce fichier contient la fonction 'main'. L'exécution du programme commence et se termine à cet endroit.
//

#include "OriginFile.h"

#include <iostream>
#include <fstream>
#include <cstdint>
#include <experimental/filesystem>

namespace fs = std::experimental::filesystem::v1;

int main(int argc, char *argv[])
{
	bool parseData = false;
	string OpjExtension = ".opj";
	bool endLine;

	if (argc < 2) {
		cout << "Usage: exportOpjFile.exe <file.opj>" << endl;
		return -1;
	}

	string input = argv[argc - 1];
	
	int folderNameIdx = input.find(OpjExtension);
	string folder = input.substr(0, folderNameIdx);

	const int dir_err = fs::create_directory(folder.c_str());
	if (dir_err == -1 && !fs::exists(folder)) {
		cout << "Error: impossible to create folder to save file" << endl;
		return -1; 
	}
	else {
		cout << "Initializing saving file in the folder: " << folder.c_str() << endl;
	}

	// Open File
	OriginFile opj(input);

	int status = opj.parse();

	cout << "Parsing status = " << status << endl;
	cout << "OPJ PROJECT \"" << input.c_str() << "\" VERSION = " << opj.version() << endl;

	cout << "number of datasets     = " << opj.datasetCount() << endl;
	cout << "number of spreadsheets = " << opj.spreadCount() << endl;
	cout << "number of matrixes     = " << opj.matrixCount() << endl;
	cout << "number of excels       = " << opj.excelCount() << endl;
	cout << "number of functions    = " << opj.functionCount() << endl;
	cout << "number of graphs       = " << opj.graphCount() << endl;
	cout << "number of notes        = " << opj.noteCount() << endl;

	cout << endl << "----------" << endl;
	
	for (int e = 0; e < 1/*opj.excelCount()*/; e++) {
		Origin::Excel excel = opj.excel(e);

		string name = excel.name;
		string label = excel.label;

		int FileNameIdx = label.find("\n");
		string FileName = label.substr(0, FileNameIdx-1);
		
		cout << "Excel " << (e + 1) << endl;
		cout << " Name: " << name.c_str() << endl;
		cout << " Label: " << FileName << endl;

		vector<Origin::SpreadSheet> sheets = excel.sheets;

		string filepath = folder + string("/") + FileName + OpjExtension;
		cout << filepath << endl; 
		ofstream outputFile(filepath.c_str(), ios::out);

		outputFile << "Name: " << label << endl; 
		outputFile << endl; 

		if (outputFile) {
			cout << "Exporting data in filename: " << filepath << endl;
		}
		else {
			cout << "Error: Unable to create file to save data" << endl;
			return -1;
		}

		for (int s = 0; s < sheets.size(); s++) {
			Origin::SpreadSheet sheet = sheets[s];

			cout << "  Spreadsheet " << (s + 1) << endl;
			cout << "   Name: " << sheet.name.c_str() << endl;
			cout << "   Label: " << sheet.label.c_str() << endl;

			if (sheet.name == "Data") parseData = true; else parseData = false;
			
			vector<Origin::SpreadColumn> columns = sheet.columns;
			int columnsCount = columns.size();
			cout << "	Columns: " << columnsCount << endl;


			for (int c = 0; c < columnsCount; c++) {
				Origin::SpreadColumn column = columns[c];
				cout << "	Column " << (c + 1) << " : " << column.name.c_str() << " / type : " << column.type << ", rows : " << sheet.maxRows << endl;
			}

				
			int dataCount = sheet.columns[0].data.size();

			if (parseData) {
				for (int d = 0; d < dataCount; d++) {
					for (int col = 0; col < columnsCount; col++) {

						string dataSeparator = "	";
						if (col == columnsCount - 1) endLine = 1; else endLine = 0;
						string sep = endLine ? "\n" : dataSeparator;

						Origin::variant value(sheet.columns[col].data[d]);

						switch (value.type()) {
						case Origin::variant::V_DOUBLE: {
							double v = value.as_double();

							if (v != _ONAN) {
								cout << v << ";" ;
								outputFile << v << sep.c_str();

							}
							else {
								cout << nan("NaN") << "; ";
								outputFile << nan("NaN") << sep.c_str();
							}

							break;
						}
						case Origin::variant::V_STRING:
							cout << value.as_string() << sep.c_str();
							break;
						default:
							cout << nan("NaN") << sep.c_str();
							break;
						}
					}
				}
			}

			outputFile << endl;
		}

		outputFile.close();
	}
}
// Exécuter le programme : Ctrl+F5 ou menu Déboguer > Exécuter sans débogage
// Déboguer le programme : F5 ou menu Déboguer > Démarrer le débogage

// Conseils pour bien démarrer : 
//   1. Utilisez la fenêtre Explorateur de solutions pour ajouter des fichiers et les gérer.
//   2. Utilisez la fenêtre Team Explorer pour vous connecter au contrôle de code source.
//   3. Utilisez la fenêtre Sortie pour voir la sortie de la génération et d'autres messages.
//   4. Utilisez la fenêtre Liste d'erreurs pour voir les erreurs.
//   5. Accédez à Projet > Ajouter un nouvel élément pour créer des fichiers de code, ou à Projet > Ajouter un élément existant pour ajouter des fichiers de code existants au projet.
//   6. Pour rouvrir ce projet plus tard, accédez à Fichier > Ouvrir > Projet et sélectionnez le fichier .sln.