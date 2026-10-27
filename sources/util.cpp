#include <iostream>
#include <iomanip>
#include <fstream>
#include <vector>
#include <sstream>
#include <algorithm>
#include <filesystem>
#include <chrono>
#include <any>
#include <xlsxio_read.h> // Cabeçalho oficial de leitura da XLSX I/O

class Xls {

    private:

    std::string strValor;

    xlsxioreader xlsLeitor;
    xlsxioreadersheet xlsFolha;

    bool bXlsLeitor;
    bool bXlsFolha;
    
    public:

    inline const std::string impCelula() const
    {
        return strValor;
    }
    inline bool okLeitor() const
    {
        return bXlsLeitor;
    }
    inline bool okFolha() const
    {
        return bXlsFolha;
    }
    inline bool lerLinha()
    {
        return (xlsxioread_sheet_next_row(xlsFolha))? true: false;
    }
    
    bool lerCelula()
    {
        char *charValor;

        if((charValor = xlsxioread_sheet_next_cell(xlsFolha)) == NULL)
            return false;

        strValor = charValor;
        xlsxioread_free(charValor);
        
        return true;
    }

    Xls(const std::string strNomeArq)
    {
        bXlsFolha = bXlsLeitor = false;

        if((xlsLeitor = xlsxioread_open(strNomeArq.c_str())) != NULL)
            bXlsLeitor = true;
        if (bXlsLeitor and (xlsFolha = xlsxioread_sheet_open(xlsLeitor, NULL, 
                  XLSXIOREAD_SKIP_EMPTY_ROWS)) != NULL)
            bXlsFolha = true;
    }
    ~Xls()
    {
        if(bXlsFolha)
            xlsxioread_sheet_close(xlsFolha);
        if(bXlsLeitor)
            xlsxioread_close(xlsLeitor);
    }
};

class FUtil {
    
    private:

    std::vector<std::vector<std::string>> vvstrTabela;
    std::vector<std::pair<int,int>> vpQuantInd;
    std::vector<float> vfPorCento;
    std::vector<std::vector<int>> vviLotofacil;

    unsigned int nTotalLin,ultimaLin;//,nTotalCol;
    
    public:

    int lerXls()
    {
        std::vector<std::string> vstrLinha;
        
        Xls lf ("../db/lotofacil.xlsx");

        if(!lf.okLeitor())
            return -1;
        else if(!lf.okFolha())
            return -2;
        else if(!lf.lerLinha())
            return -3;

        while(lf.lerLinha()) {

            while(lf.lerCelula()) {

                if(lf.impCelula() == "")
                    vstrLinha.push_back(" ");
                else 
                    vstrLinha.push_back(lf.impCelula());
            }

            vvstrTabela.push_back(vstrLinha);
            vstrLinha.clear();
        }

        return 0;
    }

    void calculaValores(int nVezes)
    {
        std::vector<int> viQuant(26,0);
        std::vector<int> viLf;

        for(auto &vT: vvstrTabela) {

            viLf.clear();
            for(int nC = 2; nC < 17; nC++)
                viLf.push_back(std::stoi(vT[nC]));

            vviLotofacil.push_back(viLf);
        }

        nTotalLin = vvstrTabela.size();
        //nTotalCol = vvstrTabela[0].size();
        ultimaLin = nTotalLin -1;
        
        for(int iL = (nTotalLin -nVezes); iL < (ultimaLin -1); iL++)
            for(int iC = 2; iC < 17; iC++)
                viQuant[std::stoi(vvstrTabela[iL][iC])]++;
        for(int nN = 1; nN <= 25; nN++)
            vpQuantInd.push_back({viQuant[nN],nN});
        std::sort(vpQuantInd.rbegin(),vpQuantInd.rend());

        for(auto &par: vpQuantInd)
            vfPorCento[par.second] = ((float) par.first /nVezes *100);
    }

    bool relatorio(int nVezes=0)
    {
        std::fstream arq;

        switch(lerXls()) {

            case -1:
                std::cout<< "Erro ao tentar abrir arquivo.\n";
                return false;
            break;
            case -2:
                std::cout<< "Erro ao tentar abrir tabela.\n";
                return false;
            break;
            case -3:
                std::cout<< "Tabela vazia.\n";
                return false;
            break;
        }

        if(nVezes <= 0)
            nVezes = nTotalLin;

        calculaValores(nVezes);

        //--> IMPRIME CONTEÚDO DE LOTOFACIL.XLSX PARA LOTOFACIL.TDB <----
        arq.open("../db/lotofacil.tdb",std::fstream::out|std::fstream::trunc);
        for(auto &t: vvstrTabela) {

            for(auto &l: t)
                arq<< l<< '#';

            arq<< std::endl;
        }
        arq.close();

        //--> IMPRIME CONTEÚDO DE QUANTLF.TDB <----
        arq.open("../db/quantlf.tdb",std::fstream::out|std::fstream::trunc);
        for(auto &par: vpQuantInd) {

            arq<< std::setw(2)<< std::setfill('0')<< par.second<< ';'
               << std::setw(2)<< std::setfill('0')<< par.first<< ';'
               << std::fixed<< std::setprecision(2)<< vfPorCento[par.second];

            if(std::find(vviLotofacil[ultimaLin].begin(),vviLotofacil[ultimaLin].end(),
                         par.second) != vviLotofacil[ultimaLin].end())
                arq<< ";*";
            
            arq<< std::endl;
        }
        arq.close();
        

        //--> IMPRIME CONTEŨDO DE GQUANTLF.TDB <----
        arq.open("../db/gquantlf.tdb",std::fstream::out|std::fstream::trunc);

        std::filesystem::path caminhoArq ("../db/gquantlf.tdb");

        auto dataCriacao = std::filesystem::last_write_time(caminhoArq);
        auto sctp = std::chrono::time_point_cast<std::chrono::system_clock::duration>(
             dataCriacao - std::filesystem::file_time_type::clock::now() + 
             std::chrono::system_clock::now());
        std::time_t tempoLegivel = std::chrono::system_clock::to_time_t(sctp);

        arq<< std::ctime(&tempoLegivel)<< "\n\t\t\tMédia: "<< nVezes<< std::endl;
        for(int nC = 0; nC < 25; nC++) {

            arq<< std::setw(2)<< std::setfill('0')<< vpQuantInd[nC].second<< " "
               << std::string(vfPorCento[vpQuantInd[nC].second] -1,'*');

            if(std::find(vviLotofacil[ultimaLin].begin(),vviLotofacil[ultimaLin].end(),
                         vpQuantInd[nC].second) != vviLotofacil[ultimaLin].end())
                arq<< 'x';
            else
                arq<< '*';

            arq<< " - "<< vfPorCento[vpQuantInd[nC].second]<< "% "<< std::endl;
        }

        arq<< "\nN. acertos 15: "<< std::setfill(' ')
           << std::setw(6)<< std::right<< vvstrTabela[ultimaLin][17]<< " "
           << vvstrTabela[ultimaLin][18]<< " "
           << vvstrTabela[ultimaLin][19]<< "\n"

           << "\nN. acertos 14: "
           << std::setw(6)<< std::right<< vvstrTabela[ultimaLin][20]<< " "
           << std::setw(10)<< std::right<< vvstrTabela[ultimaLin][21]

           << "\nN. acertos 13: "
           << std::setw(6)<< std::right<< vvstrTabela[ultimaLin][22]<< " "
           << std::setw(10)<< std::right<< vvstrTabela[ultimaLin][23]

           << "\nN. acertos 12: "
           << std::setw(6)<< std::right<< vvstrTabela[ultimaLin][24]<< " "
           << std::setw(10)<< std::right<< vvstrTabela[ultimaLin][25]

           << "\nN. acertos 11: "
           << std::setw(6)<< std::right<< vvstrTabela[ultimaLin][26]<< " "
           << std::setw(10)<< std::right<< vvstrTabela[ultimaLin][27];

        arq.close();
        
        return true;
    }

    FUtil(): vfPorCento(26,0.0)
    {

    }

} util;

int main()
{
    util.relatorio(30);
    
    return 0;
}