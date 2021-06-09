using System;
using System.Collections.Generic;
using System.Xml;
using BancoMarcasMongoDB.Models;
using BancoMarcasMongoDB.Services;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Globalization;
using System.Text;

namespace LeitorDoExtrator
{
    public static class Program
    {

        static void Main(string[] args)
        {

            string folderPath = Directory.GetCurrentDirectory() + "\\Xml";
            foreach (string file in Directory.EnumerateFiles(folderPath, "*.xml"))
            {
                //Criando a lista de processos
                List<ProcessoDadosCadastrais> listaProcessos = new List<ProcessoDadosCadastrais>();

                XmlDocument xml = new XmlDocument();
                string filePath = file;
                XElement exemploDeDados = XElement.Load(filePath);

                //criando a lista para testar o código dos processos
                List<string> despachosAntigos = new List<string> { "995", "983", "990", "994", "996", "997", "998", "992", "993", "991" };
                List<string> despachosNovos = new List<string> { "3301", "3261", "3241", "3917", "3751", "3755", "3754", "3741", "3745", "P22", "P07" };
                List<string> codigoDespachosAntigosConcessao = new List<string> { "400", "401", "403", "404", "405", "406", "450", "451", "453", "IPAS158" };

                try
                {

                    ///var t = exemploDeDados.Descendants("processo").ToList()[0.Descendants("titulares").Descendants("titular").ToList().Select(x => (string)x.Attribute("nome-razao-social")).ToList()[0];
                    for (int i = 0; i < exemploDeDados.Descendants("processo").ToList().Count; i++)
                    {

                        //instanciando as listas
                        ProcessoDadosCadastrais processoDados = new ProcessoDadosCadastrais();
                        processoDados.Titulares = new List<Titulares>();
                        processoDados.ClassesVienna = new List<ClassesVienna>();
                        processoDados.PrioridadeUnionista = new List<PrioridadeUnionista>();
                        processoDados.ClassesNice = new List<ListaClasseNice>();
                        processoDados.Despachos = new List<DespachosProcesso>();
                        processoDados.Peticoes = new List<Peticao>();
                        processoDados.Sobrestadores = new List<Sobrestadores>();
                        processoDados.ClasseNacional = new List<ClasseNacional>();
                        processoDados.DadosDeMadris = new List<DadosDeMadri>();


                        Console.WriteLine("---------------------------------------------- processo " + (i + 1) + " ----------------------------------------------");
                        var obj = exemploDeDados.Descendants("processo").ToList()[i];

                        if (obj.Attribute("numero") != null)
                        {
                            string numero_ = obj.Attribute("numero").Value;
                            Console.WriteLine("NUMERO DO PROCESSO: " + numero_);

                            processoDados.NumeroProcesso = numero_;
                        }

                        if (obj.Attribute("data-deposito") != null)
                        {
                            string dataDep_ = obj.Attribute("data-deposito").Value;
                            Console.WriteLine("Data de depósito: " + dataDep_);

                            DateTime dataDep_DateTime = Convert.ToDateTime(dataDep_);
                            processoDados.DataDeposito = dataDep_DateTime;
                        }

                        //if (obj.Attribute("data-concessao") != null)
                        //{
                        //    string dataConce_ = obj.Attribute("data-concessao").Value;
                        //    Console.WriteLine("Data de concessão: " + dataConce_);

                        //    DateTime dataConce_DateTime = Convert.ToDateTime(dataConce_);
                        //    processoDados.DataConcessao = dataConce_DateTime;
                        //}
                        //if (obj.Attribute("data-vigencia") != null)
                        //{
                        //    string dataVigen_ = obj.Attribute("data-vigencia").Value;
                        //    Console.WriteLine("Data de vingência: " + dataVigen_);

                        //    DateTime dataVigen_DateTime = Convert.ToDateTime(dataVigen_);
                        //    processoDados.DataVigencia = dataVigen_DateTime;
                        //}

                        if (obj.Descendants("procurador").ToList().Count > 0)
                        {

                            var procurador_ = obj.Descendants("procurador").ToList()[0].Value;
                            Console.WriteLine("Procurador: " + procurador_);

                            processoDados.NomeProcurador = procurador_;
                        }
                        if (obj.Descendants("apostila").ToList().Count > 0)
                        {
                            if (obj.Descendants("apostila").ToList()[0].Value != "")
                            {
                                var apostila_ = obj.Descendants("apostila").ToList()[0].Value;
                                Console.WriteLine("Apostila: " + apostila_);

                                processoDados.Apostila = apostila_;
                            }

                        }

                        if (obj.Descendants("situacao").ToList().Count > 0)
                        {
                            if (obj.Descendants("situacao").ToList()[0].Value != "")
                            {
                                var situacao_ = obj.Descendants("situacao").ToList()[0].Value;
                                Console.WriteLine("Situação: " + situacao_);

                                processoDados.Situacao = situacao_;
                            }
                        }

                        if (obj.Descendants("marca").ToList().Count > 0)
                        {
                            var marca_ = obj.Descendants("marca").ToList()[0];
                            string apresentacao_ = marca_.Attribute("apresentacao").Value;
                            string natureza_ = marca_.Attribute("natureza").Value;
                            natureza_ = natureza_.Replace("\n", " ");

                            Console.WriteLine("Apresentação: " + apresentacao_);
                            Console.WriteLine("Natureza: " + natureza_);
                            processoDados.ApresentacaoMarca = apresentacao_;
                            processoDados.NaturezaMarca = natureza_;
                            
                        }
                        if (obj.Descendants("marca").Descendants("nome").ToList().Count > 0)
                        {
                            var nome_ = obj.Descendants("marca").Descendants("nome").ToList()[0].Value;
                            Console.WriteLine("Nome: " + nome_);

                            processoDados.NomeMarca = nome_;
                        }


                        if (obj.Descendants("marca").Descendants("traducao").ToList().Count > 0)
                        {

                            if (obj.Descendants("marca").Descendants("traducao").ToList()[0].Value != "")
                            {
                                var traducao_ = obj.Descendants("marca").Descendants("traducao").ToList()[0].Value;
                                Console.WriteLine("Tradução: " + traducao_);

                                processoDados.TraducaoMarca = traducao_;
                            }
                        }

                        var titulares_ = obj.Descendants("titulares").Descendants("titular").ToList();
                        foreach (var tit in titulares_)
                        {
                            Titulares titulares = new Titulares();

                            string titular = tit.Attribute("nome-razao-social").Value;
                            Console.WriteLine("Nome - Razão social: " + titular);

                            titulares.NomeTitular = titular;
                            if (tit.Attribute("pais").Value != "")
                            {
                                string titPais = tit.Attribute("pais").Value;
                                Console.WriteLine("País: " + titPais);

                                titulares.PaisTitular = titPais;
                            }
                            if (tit.Attribute("uf").Value != "")
                            {
                                string titUf = tit.Attribute("uf").Value;
                                Console.WriteLine("UF: " + titUf);

                                titulares.UfTitular = titUf;
                            }

                            processoDados.Titulares.Add(titulares);
                        }

                        var dadosDeMadri_ = obj.Descendants("dados-de-madri").ToList(); //essa node não aparece nos dois processos, preciso arrumar uma maneira de tratar isso
                        if (dadosDeMadri_.Count > 0) //so buscará se existir esse node
                        {
                            DadosDeMadri dadosDeMadri = new DadosDeMadri();
                            var ddM = obj.Descendants("dados-de-madri").ToList()[0];

                            if (ddM.Attribute("inscricao_internacional") != null)
                            {
                                var inscInter_ = ddM.Attribute("inscricao_internacional").Value;
                                Console.WriteLine("Inscrição Internacional: " + inscInter_);

                                dadosDeMadri.NumeroInscricaoInternacional = inscInter_;
                            }
                            if (ddM.Attribute("data-recebimento-inpi") != null)
                            {
                                var dataRecebInpi_ = ddM.Attribute("data-recebimento-inpi").Value;
                                Console.WriteLine("Data de recebimento do Inpi: " + dataRecebInpi_);

                                DateTime dataRecebimentoDateTime = Convert.ToDateTime(dataRecebInpi_);
                                dadosDeMadri.DataRecebimentoInpi = dataRecebimentoDateTime;
                            }
                            processoDados.DadosDeMadris.Add(dadosDeMadri);

                            var listaNice = obj.Descendants("lista-classe-nice").Descendants("classe-nice").ToList();
                            if (listaNice.Count > 0)
                            {
                                foreach (var nic in listaNice)
                                {
                                    ListaClasseNice listaClasseNice = new ListaClasseNice();

                                    string cod = nic.Attribute("codigo").Value;
                                    Console.WriteLine("Lista de classe nice, código: " + cod);

                                    listaClasseNice.CodigoClasseNice = cod;

                                    if (nic.Descendants("edicao").ToList().Count > 0)
                                    {
                                        var edicao = nic.Descendants("edicao").ToList()[0].Value;
                                        Console.WriteLine("Edição: " + edicao);

                                        listaClasseNice.EdicaoClasseNice = edicao;
                                    }

                                    var especi = nic.Descendants("especificacao").ToList()[0].Value;
                                    Console.WriteLine("Especificação: " + especi);

                                    listaClasseNice.EspecificacaoTraducaoIngles = especi;

                                    if (nic.Descendants("traducao-especificacao").ToList().Count > 0)
                                    {
                                        var especiTraducao = nic.Descendants("traducao-especificacao").ToList()[0].Value;
                                        Console.WriteLine("Tradução da Especificação: " + especiTraducao);

                                        listaClasseNice.EspecificacaoClasseNice = especiTraducao;
                                    }
                                    processoDados.ClassesNice.Add(listaClasseNice);
                                }
                            }
                        }
                        else
                        {
                            var listaNice = obj.Descendants("lista-classe-nice").Descendants("classe-nice").ToList();
                            if (listaNice.Count > 0)
                            {
                                foreach (var nic in listaNice)
                                {
                                    ListaClasseNice listaClasseNice = new ListaClasseNice();

                                    string cod = nic.Attribute("codigo").Value;
                                    Console.WriteLine("Lista de classe nice, código: " + cod);

                                    listaClasseNice.CodigoClasseNice = cod;

                                    if (nic.Descendants("edicao").ToList().Count > 0)
                                    {
                                        var edicao = nic.Descendants("edicao").ToList()[0].Value;
                                        Console.WriteLine("Edição: " + edicao);

                                        listaClasseNice.EdicaoClasseNice = edicao;
                                    }

                                    var especi = nic.Descendants("especificacao").ToList()[0].Value;
                                    Console.WriteLine("Especificação: " + especi);

                                    listaClasseNice.EspecificacaoClasseNice = especi;

                                    if (nic.Descendants("traducao-especificacao").ToList().Count > 0)
                                    {
                                        var especiTraducao = nic.Descendants("traducao-especificacao").ToList()[0].Value;
                                        Console.WriteLine("Tradução da Especificação: " + especiTraducao);

                                        listaClasseNice.EspecificacaoTraducaoIngles = especiTraducao;
                                    }
                                    processoDados.ClassesNice.Add(listaClasseNice);
                                }
                            }
                        }

                        if (obj.Descendants("prioridade-unionista").Descendants("prioridade").ToList().Count > 0)
                        {
                            var prioriUni_ = obj.Descendants("prioridade-unionista").Descendants("prioridade").ToList();
                            foreach (var priori in prioriUni_)
                            {
                                PrioridadeUnionista prioridadeUnionista = new PrioridadeUnionista();

                                var numeroPriori = priori.Attribute("numero").Value;
                                var dataPriori = priori.Attribute("data").Value;
                                var paisPriori = priori.Attribute("pais").Value;

                                Console.WriteLine("Número prioridade: " + numeroPriori);
                                Console.WriteLine("País: " + dataPriori);
                                Console.WriteLine("Data: " + paisPriori);

                                DateTime DateTimeDataPriori = Convert.ToDateTime(dataPriori);

                                prioridadeUnionista.NumeroPrioridadeUnionista = numeroPriori;
                                prioridadeUnionista.DataPrioridadeUnionista = DateTimeDataPriori;
                                prioridadeUnionista.PaisPrioridadeUnionista = paisPriori;
                                processoDados.PrioridadeUnionista.Add(prioridadeUnionista);

                            }

                        }
                        var classesViena = obj.Descendants("classes-vienna").Descendants("classe-vienna").ToList();
                        foreach (var vie in classesViena)
                        {
                            ClassesVienna classes = new ClassesVienna();
                            string edic = vie.Attribute("edicao").Value;
                            string cod = vie.Attribute("codigo").Value;
                            Console.WriteLine("Classe vienna");
                            Console.WriteLine("Edição: " + edic);
                            Console.WriteLine("Código: " + cod);

                            classes.EdicaoClasseVienna = edic;
                            classes.CodigoClasseVienna = cod;

                            processoDados.ClassesVienna.Add(classes);
                        }

                        var classeNacional = obj.Descendants("classe-nacional").ToList();
                        if (classeNacional.Count > 0)
                        {
                            //ClasseNacional classeNacional1 = new ClasseNacional();

                            //var classeNacio = obj.Descendants("classe-nacional").ToList()[0];
                            //string codNaci = classeNacio.Attribute("codigo").Value;
                            //Console.WriteLine("Código: " + codNaci);

                            //classeNacional1.CodigoClasseNacional = codNaci;

                            var subNacional = classeNacional.Descendants("sub-classes-nacional").ToList();
                            if (subNacional.Count > 0)
                            {
                                var subClasses = subNacional.Descendants("sub-classe-nacional").ToList();
                                foreach (var sbC in subClasses)
                                {
                                    ClasseNacional classeNacional1 = new ClasseNacional();

                                    var classeNacio = obj.Descendants("classe-nacional").ToList()[0];
                                    string codNaci = classeNacio.Attribute("codigo").Value;
                                    Console.WriteLine("Código: " + codNaci);

                                    classeNacional1.CodigoClasseNacional = codNaci;

                                    string subCl = sbC.Attribute("codigo").Value;
                                    string subN = sbC.Value;

                                    Console.Write("Código " + subCl + ": ");
                                    Console.WriteLine(subN);

                                    classeNacional1.CodigoSubClasseNacional = subCl;
                                    classeNacional1.Especificacao = subN;
                                    processoDados.ClasseNacional.Add(classeNacional1);
                                    
                                }
                            }
                        }

                        //var listaNice = obj.Descendants("lista-classe-nice").Descendants("classe-nice").ToList();
                        //if (listaNice.Count > 0)
                        //{
                        //    foreach (var nic in listaNice)
                        //    {
                        //        ListaClasseNice listaClasseNice = new ListaClasseNice();

                        //        string cod = nic.Attribute("codigo").Value;
                        //        Console.WriteLine("Lista de classe nice, código: " + cod);

                        //        listaClasseNice.CodigoClasseNice = cod;

                        //        if (nic.Descendants("edicao").ToList().Count > 0)
                        //        {
                        //            var edicao = nic.Descendants("edicao").ToList()[0].Value;
                        //            Console.WriteLine("Edição: " + edicao);

                        //            listaClasseNice.EdicaoClasseNice = edicao;
                        //        }

                        //        var especi = nic.Descendants("especificacao").ToList()[0].Value;
                        //        Console.WriteLine("Especificação: " + especi);

                        //        listaClasseNice.EspecificacaoClasseNice = especi;

                        //        if (nic.Descendants("traducao-especificacao").ToList().Count > 0)
                        //        {
                        //            var especiTraducao = nic.Descendants("traducao-especificacao").ToList()[0].Value;
                        //            Console.WriteLine("Tradução da Especificação: " + especiTraducao);

                        //            listaClasseNice.EspecificacaoTraducao = especiTraducao;
                        //        }
                        //        processoDados.ClassesNice.Add(listaClasseNice);
                        //    }
                        //}

                        var peticao = obj.Descendants("peticoes").Descendants("peticao").ToList();
                        if (peticao != null && peticao.Count > 0)
                        {
                            foreach (var pet in peticao)
                            {
                                Peticao peticoes = new Peticao();

                                string numPet = pet.Attribute("numero").Value;
                                string dtPet = pet.Attribute("data").Value;
                                string codSerPet = pet.Attribute("codigoServico").Value;
                                var rq = pet.Descendants("requerente").ToList()[0];
                                string rqt = rq.Attribute("nome-razao-social").Value;

                                Console.WriteLine("Número da petição: " + numPet);
                                Console.WriteLine("Data da petição: " + dtPet);
                                Console.WriteLine("Código de serviço: " + codSerPet);
                                Console.WriteLine("Nome - Razão social do requerente: " + rqt);

                                DateTime dateTimeDtPet = Convert.ToDateTime(dtPet);

                                peticoes.PeticaoProtocolo = numPet;
                                peticoes.PeticaoData = dateTimeDtPet;
                                peticoes.PeticaoServico = codSerPet;
                                peticoes.PeticaoRequerente = rqt;

                                processoDados.Peticoes.Add(peticoes);
                            }
                        }



                        var despacho = obj.Descendants("despachos").Descendants("despacho").ToList().OrderBy(x => int.Parse(x.Attribute("rpi").Value.Trim()));
                        foreach (var desp in despacho)
                        {
                            DespachosProcesso despachos = new DespachosProcesso();

                            if (desp.Attribute("rpi") != null)
                            {

                                string rpi = desp.Attribute("rpi").Value;

                                Console.WriteLine("Rpi: " + rpi);

                                despachos.Rpi = rpi;

                            }
                            if (desp.Attribute("data-rpi") != null)
                            {

                                string dtRpi = desp.Attribute("data-rpi").Value;
                                Console.WriteLine("Data rpi: " + dtRpi);

                                DateTime dateTimeDtRpi = Convert.ToDateTime(dtRpi);

                                despachos.DataRpiDespacho = dateTimeDtRpi;
                            }
                            if (desp.Attribute("despacho-rpi") != null)
                            {
                                string despRpi = desp.Attribute("despacho-rpi").Value;
                                Console.WriteLine("Despacho Rpi: " + despRpi);

                                if (despRpi.Length != 3 && despRpi.Length != 7)
                                {
                                    despachos.RpiNomeDespacho = despRpi;
                                }
                                else
                                {
                                    despachos.RpiCodigoDespacho = despRpi;
                                }

                                //Lógica para concessão
                                if (codigoDespachosAntigosConcessao.Exists(x => x == despRpi) || RemoveDiacritics(despRpi).ToLower() == "concessao de registro")
                                {
                                    processoDados.DataConcessao = despachos.DataRpiDespacho;
                                    DateTime dataRegistro = processoDados.DataConcessao;
                                    processoDados.DataRegistro = dataRegistro;

                                    DateTime dataProrrogacaoOrdinarioInicial = processoDados.DataConcessao.AddYears(10).AddYears(-1).AddDays(1);
                                    processoDados.ProrrogacaoDataOrdinarioInicial = dataProrrogacaoOrdinarioInicial;

                                    DateTime dataProrrogacaoOrdinarioFim = processoDados.DataConcessao.AddYears(10);
                                    processoDados.ProrrogacaoDataOrdinarioFinal = dataProrrogacaoOrdinarioFim;

                                    DateTime dataVigencia = dataProrrogacaoOrdinarioFim;
                                    processoDados.DataVigencia = dataVigencia;

                                    DateTime dataExtraordinarioInicio = processoDados.DataConcessao.AddYears(10).AddDays(1);
                                    processoDados.ProrrogacaoDataExtraOrdinarioInicial = dataExtraordinarioInicio;

                                    DateTime dataExtraordinarioFinal = dataProrrogacaoOrdinarioFim.AddMonths(6);
                                    processoDados.ProrrogacaoDataExtraOrdinarioFinal = dataExtraordinarioFinal;
                                }
                            }

                            if (desp.Descendants("texto-complementar").ToList().Count > 0)
                            {
                                var textComp = desp.Descendants("texto-complementar").ToList()[0].Value;
                                Console.WriteLine("Texto complementar: " + textComp);

                                despachos.RpiComplementoDespacho = textComp;

                            }

                            var prot = desp.Descendants("protocolo").ToList();
                            if (prot != null && prot.Count > 0)
                            {
                                var proto = desp.Descendants("protocolo").ToList()[0];
                                string numProto = proto.Attribute("numero").Value;
                                if (proto.Attribute("data") != null)
                                {
                                    string dtProto = proto.Attribute("data").Value;
                                    Console.WriteLine("Data: " + dtProto);
                                    if (!String.IsNullOrEmpty(dtProto))
                                    {
                                        DateTime dateTimeDtProto = Convert.ToDateTime(dtProto);
                                        despachos.RpiProtocoloData = dateTimeDtProto;
                                    }

                                }
                                string codServico = proto.Attribute("codigoServico").Value;
                                Console.WriteLine("Número: " + numProto);
                                Console.WriteLine("Código do serviço: " + codServico);


                                despachos.RpiProtocoloNumero = numProto;
                                despachos.RpiCodigoServicoProtocolo = codServico;


                                if (prot.Descendants("requerente").ToList().Count > 0)
                                {
                                    var requerente = prot.Descendants("requerente").ToList()[0];
                                    if (requerente.Attribute("nome-razao-social") != null)
                                    {
                                        string nomeRequerente = requerente.Attribute("nome-razao-social").Value;

                                        Console.WriteLine("Requerente: " + nomeRequerente);
                                        despachos.RpiRequerente = nomeRequerente;
                                    }


                                    if (requerente.Descendants("procurador").ToList().Count > 0 && requerente.Descendants("procurador") != null)
                                    {
                                        var procurador = requerente.Descendants("procurador").ToList()[0].Value;

                                        despachos.RpiProcurador = procurador;
                                    }

                                    if (requerente.Descendants("cessionario").ToList().Count > 0 && requerente.Descendants("cessionario").ToList() != null)
                                    {
                                        var cessionario = requerente.Descendants("cessionario").ToList();
                                        foreach (var item in cessionario)
                                        {
                                            string nomeCessionario = item.Attribute("nome-razao-social").Value;

                                            despachos.NomeCessionario = nomeCessionario;
                                        }
                                    }
                                }
                            }

                            if (processoDados.DataRegistro != DateTime.MinValue)
                            {
                                //Lógica para realizar a prorrogação
                                if (despachosNovos.Exists(x => x == despachos.RpiCodigoServicoProtocolo && despachos.RpiNomeDespacho.Equals("Deferimento da petição", StringComparison.OrdinalIgnoreCase)) || despachosAntigos.Exists(x => x == despachos.RpiCodigoDespacho))
                                {

                                    DateTime dataRegistro = processoDados.DataVigencia;
                                    processoDados.DataRegistro = dataRegistro;

                                    DateTime dataVigencia = processoDados.DataRegistro.AddYears(10);
                                    processoDados.DataVigencia = dataVigencia;

                                    DateTime dataProrrogacaoOrdinarioInicial = processoDados.DataRegistro.AddYears(10).AddYears(-1).AddDays(1);
                                    processoDados.ProrrogacaoDataOrdinarioInicial = dataProrrogacaoOrdinarioInicial;

                                    DateTime dataProrrogacaoOrdinarioFim = processoDados.DataRegistro.AddYears(10);
                                    processoDados.ProrrogacaoDataOrdinarioFinal = dataProrrogacaoOrdinarioFim;

                                    DateTime dataExtraordinarioInicio = processoDados.DataRegistro.AddYears(10).AddDays(1);
                                    processoDados.ProrrogacaoDataExtraOrdinarioInicial = dataExtraordinarioInicio;

                                    DateTime dataExtraordinarioFinal = dataProrrogacaoOrdinarioFim.AddMonths(6);
                                    processoDados.ProrrogacaoDataExtraOrdinarioFinal = dataExtraordinarioFinal;

                                }
                            }

                            processoDados.Despachos.Add(despachos);
                            //listaProcessos.Add(processoDados);
                        }
                        listaProcessos.Add(processoDados);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("Ocorreu um erro: " + e.Message);
                }

                string filename = null;
                filename = Path.GetFileName(file);


                string caminho = Directory.GetCurrentDirectory() + "\\Log\\Log_" + filename + ".txt";
                string caminhoAlterado = Directory.GetCurrentDirectory() + "\\XML armazenados\\armazenado_" + filename;

                using (StreamWriter sw = File.AppendText(caminho))
                {
                    foreach (var proc in listaProcessos)
                    {
                        try
                        {

                            ProcessoDadosCadastrais existe = ProcessoServices.GetProcessoByNumero(proc.NumeroProcesso);
                            if (existe == null)
                            {
                                proc.DataImportacao = DateTime.Now;
                                ProcessoServices.InsertProcessoDados(proc);
                                Console.WriteLine("Processo " + proc.NumeroProcesso + " gravado com sucesso!");

                                sw.WriteLine("Processo " + proc.NumeroProcesso + " gravado com sucesso! em " + proc.DataImportacao);
                                sw.WriteLine("------------------------------------------------");
                                sw.WriteLine();

                            }
                            else
                            {
                                proc.DataImportacao = DateTime.Now;
                                ProcessoServices.UpdateProcessoDados(proc, existe);
                                Console.WriteLine("Processo " + proc.NumeroProcesso + " foi alterado com sucesso! em " + proc.DataImportacao);

                                sw.WriteLine("Processo " + proc.NumeroProcesso + " alterado com sucesso!");
                                sw.WriteLine("------------------------------------------------");
                                sw.WriteLine();

                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Ocorreu um erro: " + e.Message);

                            sw.WriteLine("Processo " + proc.NumeroProcesso + " com erro. Erro: " + e.Message);
                            sw.WriteLine("------------------------------------------------");
                            sw.WriteLine();
                        }
                    }
                    System.IO.File.Move(file, caminhoAlterado, true);
                    Console.WriteLine("Caminho alterado");
                }
            }
        }
        public static string RemoveDiacritics(this string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            text = text.Normalize(NormalizationForm.FormD);
            var chars = text.Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray();
            return new string(chars).Normalize(NormalizationForm.FormC);
        }
    }
}
