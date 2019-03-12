// exportOpjFile.cpp : Ce fichier contient la fonction 'main'. L'exécution du programme commence et se termine à cet endroit.
//

#include "OriginFile.h"
#include <iostream>

int main(int argc, char *argv[])
{
	bool parseData = false;
	char OpjExtension = '.opj';

	if (argc < 2) {
		cout << "Usage: exportOpjFile.exe <file.opj>" << endl;
		return -1;
	}

	string input = argv[argc - 1];
	
	int folderNameIdx = input.find(OpjExtension);
	string folder = input.substr(1, folderNameIdx);

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

		int FileNameIdx = label.find(EOL);
		string FileName = label.substr(1, FileNameIdx);
		
		cout << "Excel " << (e + 1) << endl;
		cout << " Name: " << name.c_str() << endl;
		cout << " Label: " << label.c_str() << endl;

		vector<Origin::SpreadSheet> sheets = excel.sheets;

		for (int s = 0; s < sheets.size(); s++) {
			Origin::SpreadSheet sheet = sheets[s];

			cout << "Spreadsheet " << (s + 1) << endl;
			cout << " Name: " << sheet.name.c_str() << endl;
			cout << " Label: " << sheet.label.c_str() << endl;

			if (sheet.name == "Data") parseData = true; else parseData = false;
			
			vector<Origin::SpreadColumn> columns = sheet.columns;
			int columnsCount = columns.size();
			cout << "	Columns: " << columnsCount << endl;


			for (int c = 0; c < columnsCount; c++) {
				Origin::SpreadColumn column = columns[c];
				cout << "	Column " << (c + 1) << " : " << column.name.c_str() << " / type : " << column.type << ", rows : " << sheet.maxRows << endl;

				int dataCount = column.data.size();

				if (parseData) {
					for (int d = 0; d < dataCount; d++) {
						Origin::variant value(column.data[d]);

						switch (value.type()) {
						case Origin::variant::V_DOUBLE: {
							double v = value.as_double();

							if (v != _ONAN) {
								cout << v << "; ";
							}
							else {
								cout << nan("NaN") << "; ";
							}

							break;
						}
						case Origin::variant::V_STRING:
							cout << value.as_string() << ",";
							break;
						default:
							cout << nan("NaN") << ",";
							break;
						}
					}

					cout << endl;
				}

				
			}


			
		}
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