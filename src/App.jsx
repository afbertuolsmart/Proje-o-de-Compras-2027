import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

import {
  Filter,
  RotateCcw,
  Package,
  Palette,
  Layers,
  TrendingUp,
  ChevronDown,
  Check,
  Target,
  Ruler,
  BarChart3,
  Upload,
  FileSpreadsheet
} from 'lucide-react';

import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  LabelList,
  LineChart,
  Line
} from 'recharts';

import './styles.css';


/* =========================================================
   FUNÇÕES
========================================================= */

function normalizarTexto(valor) {
  return String(valor ?? '')
    .trim()
    .replace(/\s+/g, ' ');
}

/*
   Padronização gerencial de materiais.
   NEW BASE VITÓRIA WZ 1.1 e
   NEW BASE VITÓRIA WZ 1.1 (COM MARCA D'ÁGUA)
   são tratados como um único item.
*/
function normalizarMaterial(valor) {
  let material = normalizarTexto(valor);

  material = material
    .replace(/\s*\(\s*COM\s+MARCA\s+D['’]ÁGUA\s*\)\s*$/i, '')
    .replace(/\s*\(\s*COM\s+MARCA\s+D['’]AGUA\s*\)\s*$/i, '')
    .replace(/\s+COM\s+MARCA\s+D['’]ÁGUA\s*$/i, '')
    .replace(/\s+COM\s+MARCA\s+D['’]AGUA\s*$/i, '')
    .trim();

  // Trata as duas descrições como o mesmo material
  const comparacao = material
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

  if (comparacao === 'NEW BASE VITORIA WZ 1.1') {
    return 'NEW BASE VITÓRIA WZ 1.1';
  }

  return material;
}

function numero(valor) {

  if (typeof valor === 'number') {
    return valor;
  }

  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {
    return 0;
  }

  let texto = String(valor)
    .trim();

  /*
    Trata números brasileiros e números
    já convertidos pelo Excel.
  */

  if (
    texto.includes('.') &&
    texto.includes(',')
  ) {

    texto = texto
      .replace(/\./g, '')
      .replace(',', '.');

  } else if (texto.includes(',')) {

    texto = texto.replace(',', '.');

  }

  const resultado = Number(texto);

  return Number.isFinite(resultado)
    ? resultado
    : 0;
}

function formatarNumero(valor, casas = 0) {

  return Number(valor || 0).toLocaleString(
    'pt-BR',
    {
      minimumFractionDigits: casas,
      maximumFractionDigits: casas
    }
  );

}

function formatarPercentual(valor) {

  return `${Number(valor || 0).toLocaleString(
    'pt-BR',
    {
      minimumFractionDigits: 1,
      maximumFractionDigits: 1
    }
  )}%`;

}


/* =========================================================
   MULTI SELECT
========================================================= */

function MultiSelect({
  label,
  icon: Icon,
  options,
  selected,
  setSelected,
  placeholder
}) {

  const [aberto, setAberto] =
    useState(false);

  const ref =
    useRef(null);

  useEffect(() => {

    function fechar(event) {

      if (
        ref.current &&
        !ref.current.contains(event.target)
      ) {
        setAberto(false);
      }

    }

    document.addEventListener(
      'mousedown',
      fechar
    );

    return () => {

      document.removeEventListener(
        'mousedown',
        fechar
      );

    };

  }, []);


  function alternar(item) {

    if (selected.includes(item)) {

      setSelected(
        selected.filter(
          x => x !== item
        )
      );

    } else {

      setSelected([
        ...selected,
        item
      ]);

    }

  }


  function selecionarTodos() {

    setSelected([...options]);

  }


  function desmarcarTodos() {

    setSelected([]);

  }


  let texto =
    placeholder;

  if (selected.length === 1) {

    texto =
      selected[0];

  } else if (selected.length > 1) {

    texto =
      `${selected.length} selecionados`;

  }


  return (

    <div
      className="filter-group"
      ref={ref}
    >

      <label>

        {Icon && (
          <Icon size={14} />
        )}

        {label}

      </label>


      <div className="multi-select">

        <button
          type="button"
          className="multi-select-button"
          onClick={() =>
            setAberto(!aberto)
          }
        >

          <span>
            {texto}
          </span>

          <ChevronDown
            size={16}
            className={
              aberto
                ? 'rotate'
                : ''
            }
          />

        </button>


        {aberto && (

          <div className="multi-select-menu">

            <div className="multi-select-actions">

              <button
                type="button"
                onClick={
                  selecionarTodos
                }
              >

                <Check size={13} />

                Todos

              </button>


              <button
                type="button"
                onClick={
                  desmarcarTodos
                }
              >
                Limpar
              </button>

            </div>


            <div className="multi-select-options">

              {options.map(item => {

                const marcado =
                  selected.includes(item);

                return (

                  <label
                    key={item}
                    className="multi-select-option"
                  >

                    <input
                      type="checkbox"
                      checked={marcado}
                      onChange={() =>
                        alternar(item)
                      }
                    />

                    <span className="custom-checkbox">

                      {marcado && (
                        <Check size={12} />
                      )}

                    </span>

                    <span>
                      {item}
                    </span>

                  </label>

                );

              })}

            </div>

          </div>

        )}

      </div>

    </div>

  );

}


/* =========================================================
   LOCALIZA NACIONAL
========================================================= */

function encontrarTabelaVendaGeralConsolidada(matriz) {
  // Procura TODAS as ocorrências da tabela e escolhe a que realmente
  // contém 12 meses de 2027 + valores numéricos em NACIONAL.
  const candidatos = [];

  for (let linhaTitulo = 0; linhaTitulo < matriz.length; linhaTitulo++) {
    const rowTitulo = matriz[linhaTitulo] || [];
    const temTitulo = rowTitulo.some(
      valor => normalizarTexto(valor).toUpperCase() === 'VENDA GERAL CONSOLIDADA'
    );

    if (!temTitulo) continue;

    for (let linhaCabecalho = linhaTitulo + 1; linhaCabecalho <= Math.min(matriz.length - 1, linhaTitulo + 12); linhaCabecalho++) {
      const row = matriz[linhaCabecalho] || [];

      for (let colunaNacional = 0; colunaNacional < row.length; colunaNacional++) {
        if (normalizarTexto(row[colunaNacional]).toUpperCase() !== 'NACIONAL') continue;

        // Testa cada coluna como possível coluna de mês.
        for (let colunaMes = 0; colunaMes < row.length; colunaMes++) {
          if (colunaMes === colunaNacional) continue;

          let encontrados = 0;
          let soma = 0;
          let primeiroMes = null;

          for (let i = linhaCabecalho + 1; i < Math.min(matriz.length, linhaCabecalho + 20); i++) {
            const linha = matriz[i] || [];
            const mes = converterMes(linha[colunaMes]);
            const nacional = numero(linha[colunaNacional]);

            if (mes && mes.startsWith('2027-') && nacional > 0) {
              encontrados++;
              soma += nacional;
              if (!primeiroMes) primeiroMes = mes;
            }
          }

          if (encontrados >= 10 && soma > 0 && primeiroMes) {
            candidatos.push({
              linhaTitulo,
              linhaCabecalho,
              colunaMes,
              colunaNacional,
              encontrados,
              soma
            });
          }
        }
      }
    }
  }

  if (!candidatos.length) return null;

  // Prioriza a candidata com 12 meses e maior soma de NACIONAL.
  candidatos.sort((a, b) => {
    if (b.encontrados !== a.encontrados) return b.encontrados - a.encontrados;
    return b.soma - a.soma;
  });

  return candidatos[0];
}


/* =========================================================
   CONVERTE MÊS
========================================================= */

function converterMes(valor) {
  if (valor === null || valor === undefined || valor === '') return null;

  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    return `${valor.getFullYear()}-${String(valor.getMonth() + 1).padStart(2, '0')}`;
  }

  if (typeof valor === 'number') {
    const data = XLSX.SSF.parse_date_code(valor);
    if (data && data.y >= 2000 && data.y <= 2100 && data.m >= 1 && data.m <= 12) {
      return `${data.y}-${String(data.m).padStart(2, '0')}`;
    }
  }

  const texto = normalizarTexto(valor).toLowerCase();
  const meses = {
    jan: '01', janeiro: '01', fev: '02', fevereiro: '02',
    mar: '03', marco: '03', março: '03', abr: '04', abril: '04',
    mai: '05', maio: '05', jun: '06', junho: '06', jul: '07', julho: '07',
    ago: '08', agosto: '08', set: '09', setembro: '09', out: '10', outubro: '10',
    nov: '11', novembro: '11', dez: '12', dezembro: '12'
  };

  for (const chave in meses) {
    if (texto.includes(chave)) {
      const ano = texto.match(/20\d{2}/)?.[0] || '2027';
      return `${ano}-${meses[chave]}`;
    }
  }

  const match = texto.match(/^(\d{4})-(\d{1,2})$/);
  if (match) return `${match[1]}-${String(match[2]).padStart(2, '0')}`;

  return null;
}


/* =========================================================
   LEITOR DA PLANILHA DE CONSUMO
   ========================================================= */

const STORAGE_KEY_CONSUMO = 'analise-consumo-dados-v1';
const STORAGE_KEY_CONSUMO_META = 'analise-consumo-meta-v1';
const STORAGE_KEY_PROJECAO = 'analise-projecao-2027-v1';
const STORAGE_KEY_PROJECAO_META = 'analise-projecao-2027-meta-v1';

function encontrarColuna(linha, nomes) {
  const chaves = Object.keys(linha || {});
  return chaves.find(chave => {
    const normalizada = normalizarTexto(chave)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    return nomes.some(nome => normalizada === nome);
  });
}

function analisarPlanilhaConsumo(buffer) {
  const workbook = XLSX.read(buffer, {
    type: 'array',
    cellDates: true
  });

  const nomePlanilha = workbook.SheetNames.find(nome => {
    const normalizada = normalizarTexto(nome)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
    return normalizada === 'material e cor por mes';
  });

  if (!nomePlanilha) {
    throw new Error(
      'A aba "Material e Cor por Mes" não foi encontrada. Use uma planilha no mesmo formato do arquivo atual.'
    );
  }

  const worksheet = workbook.Sheets[nomePlanilha];
  const linhas = XLSX.utils.sheet_to_json(worksheet, {
    defval: '',
    raw: false
  });

  if (!linhas.length) {
    throw new Error('A aba "Material e Cor por Mes" não possui dados.');
  }

  const primeira = linhas[0];
  const colunaMes = encontrarColuna(primeira, ['mes']);
  const colunaCor = encontrarColuna(primeira, ['cor']);
  const colunaMaterial = encontrarColuna(primeira, ['material']);
  const colunaConsumo = encontrarColuna(primeira, [
    'consumo material',
    'consumo'
  ]);

  if (!colunaMes || !colunaCor || !colunaMaterial || !colunaConsumo) {
    throw new Error(
      'A planilha não possui as colunas esperadas: Mês, Cor, Material e Consumo Material.'
    );
  }

  const dados = linhas
    .map((linha, index) => ({
      id: index,
      Mes: normalizarTexto(linha[colunaMes]),
      Cor: normalizarTexto(linha[colunaCor]),
      Material: normalizarMaterial(linha[colunaMaterial]),
      Consumo: numero(linha[colunaConsumo])
    }))
    .filter(item =>
      item.Mes &&
      item.Cor &&
      item.Material &&
      Number.isFinite(item.Consumo)
    );

  if (!dados.length) {
    throw new Error(
      'Nenhuma linha válida foi encontrada na aba "Material e Cor por Mes".'
    );
  }

  return dados;
}

function normalizarDadosConsumo(dados) {
  if (!Array.isArray(dados)) return [];

  return dados
    .map((item, index) => ({
      ...item,
      id: item.id ?? index,
      Mes: normalizarTexto(item.Mes),
      Cor: normalizarTexto(item.Cor),
      Material: normalizarMaterial(item.Material),
      Consumo: numero(item.Consumo)
    }))
    .filter(item =>
      item.Mes &&
      item.Cor &&
      item.Material &&
      Number.isFinite(item.Consumo)
    );
}

function lerDadosSalvos() {
  try {
    const salvo = localStorage.getItem(STORAGE_KEY_CONSUMO);
    if (!salvo) return null;

    const dados = JSON.parse(salvo);
    const normalizados = normalizarDadosConsumo(dados);

    if (!normalizados.length) return null;

    // Migra dados antigos para o formato normalizado e mantém
    // a correção depois de atualizar a página.
    localStorage.setItem(
      STORAGE_KEY_CONSUMO,
      JSON.stringify(normalizados)
    );

    return normalizados;
  } catch (error) {
    console.warn('Não foi possível ler os dados salvos:', error);
    return null;
  }
}

function salvarDadosLocalmente(dados, meta) {
  localStorage.setItem(
    STORAGE_KEY_CONSUMO,
    JSON.stringify(dados)
  );

  localStorage.setItem(
    STORAGE_KEY_CONSUMO_META,
    JSON.stringify(meta)
  );
}

function lerMetaLocal() {
  try {
    const meta = localStorage.getItem(STORAGE_KEY_CONSUMO_META);
    return meta ? JSON.parse(meta) : null;
  } catch {
    return null;
  }
}


function lerProjecaoSalva() {
  try {
    const salvo = localStorage.getItem(STORAGE_KEY_PROJECAO);
    if (!salvo) return null;
    const dados = JSON.parse(salvo);
    return Array.isArray(dados) && dados.length ? dados : null;
  } catch { return null; }
}

function lerMetaProjecaoLocal() {
  try {
    const salvo = localStorage.getItem(STORAGE_KEY_PROJECAO_META);
    return salvo ? JSON.parse(salvo) : null;
  } catch { return null; }
}

function salvarProjecaoLocalmente(dados, meta) {
  localStorage.setItem(STORAGE_KEY_PROJECAO, JSON.stringify(dados));
  localStorage.setItem(STORAGE_KEY_PROJECAO_META, JSON.stringify(meta));
}


/* =========================================================
   APP
========================================================= */

export default function App() {


  /* =======================================================
     ESTADOS
  ======================================================= */

  const [data, setData] =
    useState([]);

  const [origemDados, setOrigemDados] =
    useState('Arquivo padrão');

  const [atualizandoDados, setAtualizandoDados] =
    useState(false);

  const [projecaoVendas, setProjecaoVendas] =
    useState([]);

  const [loading, setLoading] =
    useState(true);

  const [loadingProjecao, setLoadingProjecao] =
    useState(true);

  const [erro, setErro] =
    useState('');

  const [erroProjecao, setErroProjecao] =
    useState('');


  /* =======================================================
     FILTROS
  ======================================================= */

  const [
    coresSelecionadas,
    setCoresSelecionadas
  ] = useState([]);

  const [
    materiaisSelecionados,
    setMateriaisSelecionados
  ] = useState([]);

  /* Filtro independente da tabela de projeção por base + cor */
  const [
    materiaisProjecaoSelecionados,
    setMateriaisProjecaoSelecionados
  ] = useState([]);

  const [
    mesesSelecionados,
    setMesesSelecionados
  ] = useState([]);

  const [
    periodo,
    setPeriodo
  ] = useState(12);

  const arquivoProjecaoRef = useRef(null);

  const [origemProjecao, setOrigemProjecao] = useState('Arquivo padrão');
  const [atualizandoProjecao, setAtualizandoProjecao] = useState(false);


  /* =======================================================
     CARREGA CONSUMO
  ======================================================= */

  useEffect(() => {

    async function carregarDados() {

      try {

        setLoading(true);

        const dadosSalvos = lerDadosSalvos();
        const metaSalva = lerMetaLocal();

        if (dadosSalvos) {
          setData(dadosSalvos);
          setOrigemDados(
            metaSalva?.nomeArquivo
              ? `Importado: ${metaSalva.nomeArquivo}`
              : 'Dados importados salvos neste navegador'
          );
          return;
        }

        const resposta = await fetch('/data/Consumo.xlsx');

        if (!resposta.ok) {
          throw new Error(
            'Não foi possível carregar Consumo.xlsx'
          );
        }

        const buffer = await resposta.arrayBuffer();
        const dados = analisarPlanilhaConsumo(buffer);

        setData(dados);
        setOrigemDados('Arquivo padrão: Consumo.xlsx');

      } catch (error) {

        console.error(error);
        setErro(error.message);

      } finally {

        setLoading(false);

      }

    }

    carregarDados();

  }, []);


  /* =======================================================
     IMPORTAÇÃO MANUAL DA PROJEÇÃO 2027
  ======================================================= */

  const NOME_ARQUIVO_PROJECAO = '17-VENDAS- PROJEÇÃO_2027.xlsx';

  async function importarPlanilhaProjecao(event) {
    const arquivo = event.target.files?.[0];
    event.target.value = '';
    if (!arquivo) return;

    try {
      setAtualizandoProjecao(true);
      setErroProjecao('');

      const buffer = await arquivo.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const nomeAnalise = workbook.SheetNames.find(nome =>
        normalizarTexto(nome).toUpperCase() === 'ANALISE'
      );
      if (!nomeAnalise) throw new Error('A planilha precisa possuir a aba ANALISE.');

      const matriz = XLSX.utils.sheet_to_json(workbook.Sheets[nomeAnalise], {
        header: 1, defval: ''
      });
      const tabela = encontrarTabelaVendaGeralConsolidada(matriz);
      if (!tabela) {
        throw new Error('A tabela VENDA GERAL CONSOLIDADA com NACIONAL 2027 não foi encontrada.');
      }

      const resultados = [];
      for (let i = tabela.linhaCabecalho + 1; i < Math.min(matriz.length, tabela.linhaCabecalho + 20) && resultados.length < 12; i++) {
        const linha = matriz[i] || [];
        const mes = converterMes(linha[tabela.colunaMes]);
        const nacional = numero(linha[tabela.colunaNacional]);
        if (mes?.startsWith('2027-') && nacional > 0 && !resultados.some(item => item.Mes === mes)) {
          resultados.push({ Mes: mes, Nacional: nacional });
        }
      }
      resultados.sort((a, b) => a.Mes.localeCompare(b.Mes));
      if (resultados.length !== 12) {
        throw new Error(`Foram encontrados ${resultados.length} meses de 2027. A projeção precisa ter 12 meses.`);
      }

      const meta = {
        nomeArquivo: arquivo.name,
        dataImportacao: new Date().toISOString(),
        linhas: resultados.length,
        total: resultados.reduce((soma, item) => soma + item.Nacional, 0)
      };

      salvarProjecaoLocalmente(resultados, meta);
      setProjecaoVendas(resultados);
      setOrigemProjecao(`Importado: ${arquivo.name}`);
    } catch (error) {
      console.error(error);
      setErroProjecao(error.message);
    } finally {
      setAtualizandoProjecao(false);
    }
  }

  /* =======================================================
     CARREGA ANALISE / NACIONAL
  ======================================================= */

  useEffect(() => {

    async function carregarProjecao() {

      try {

        setLoadingProjecao(true);

        const projecaoSalva = lerProjecaoSalva();
        const metaProjecaoSalva = lerMetaProjecaoLocal();

        if (projecaoSalva) {
          setProjecaoVendas(projecaoSalva);
          setOrigemProjecao(metaProjecaoSalva?.nomeArquivo ? `Importado: ${metaProjecaoSalva.nomeArquivo}` : 'Projeção importada salva neste navegador');
          return;
        }

        const resposta =
          await fetch(
            `/data/${encodeURIComponent(NOME_ARQUIVO_PROJECAO)}`
          );

        if (!resposta.ok) {

          throw new Error(
            `Não foi possível carregar ${NOME_ARQUIVO_PROJECAO}`
          );

        }


        const buffer =
          await resposta.arrayBuffer();


        const workbook =
          XLSX.read(
            buffer,
            {
              type: 'array',
              cellDates: true
            }
          );


        /*
          ABA ANALISE
        */

        const nomeAnalise =
          workbook.SheetNames.find(
            nome =>
              normalizarTexto(nome)
                .toUpperCase() ===
              'ANALISE'
          );


        if (!nomeAnalise) {

          throw new Error(
            'A aba ANALISE não foi encontrada.'
          );

        }


        const worksheet =
          workbook.Sheets[
            nomeAnalise
          ];


        /*
          Matriz completa
        */

        const matriz =
          XLSX.utils.sheet_to_json(
            worksheet,
            {
              header: 1,
              defval: ''
            }
          );


        const tabela = encontrarTabelaVendaGeralConsolidada(matriz);

        if (!tabela) {
          throw new Error(
            'A tabela VENDA GERAL CONSOLIDADA não foi encontrada.'
          );
        }

        const resultados = [];

        // Lê somente a série 2027 encontrada dentro da tabela correta.
        for (
          let i = tabela.linhaCabecalho + 1;
          i < Math.min(matriz.length, tabela.linhaCabecalho + 20) && resultados.length < 12;
          i++
        ) {
          const linha = matriz[i] || [];
          const mes = converterMes(linha[tabela.colunaMes]);
          const nacional = numero(linha[tabela.colunaNacional]);

          if (mes && mes.startsWith('2027-') && nacional > 0) {
            if (!resultados.some(item => item.Mes === mes)) {
              resultados.push({ Mes: mes, Nacional: nacional });
            }
          }
        }

        resultados.sort((a, b) => a.Mes.localeCompare(b.Mes));

        if (resultados.length !== 12) {
          throw new Error(
            `Foram encontrados ${resultados.length} meses de 2027 na VENDA GERAL CONSOLIDADA. Esperado: 12.`
          );
        }

        setProjecaoVendas(
          resultados
        );
        setOrigemProjecao(`Arquivo padrão: ${NOME_ARQUIVO_PROJECAO}`);


        console.log(
          'NACIONAL 2027:',
          resultados
        );


      } catch (error) {

        console.error(error);

        setErroProjecao(
          error.message
        );

      } finally {

        setLoadingProjecao(false);

      }

    }


    carregarProjecao();

  }, []);


  /* =======================================================
     LISTAS
  ======================================================= */

  const meses = useMemo(
    () =>
      [
        ...new Set(
          data
            .map(
              x => x.Mes
            )
            .filter(Boolean)
        )
      ].sort(),
    [data]
  );


  const cores = useMemo(
    () =>
      [
        ...new Set(
          data
            .map(
              x => x.Cor
            )
            .filter(Boolean)
        )
      ].sort(),
    [data]
  );


  const materiais = useMemo(
    () =>
      [
        ...new Set(
          data
            .map(
              x => x.Material
            )
            .filter(Boolean)
        )
      ].sort(),
    [data]
  );


  /* =======================================================
     MESES DO PERÍODO
  ======================================================= */

  const mesesDoPeriodo =
    useMemo(() => {

      return meses.slice(
        -Number(periodo)
      );

    }, [
      meses,
      periodo
    ]);


  /* =======================================================
     DADOS FILTRADOS
  ======================================================= */

  const dadosFiltrados =
    useMemo(() => {

      let dados =
        [...data];


      if (
        mesesSelecionados.length > 0
      ) {

        dados =
          dados.filter(
            item =>
              mesesSelecionados.includes(
                item.Mes
              )
          );

      } else {

        dados =
          dados.filter(
            item =>
              mesesDoPeriodo.includes(
                item.Mes
              )
          );

      }


      if (
        coresSelecionadas.length > 0
      ) {

        dados =
          dados.filter(
            item =>
              coresSelecionadas.includes(
                item.Cor
              )
          );

      }


      if (
        materiaisSelecionados.length > 0
      ) {

        dados =
          dados.filter(
            item =>
              materiaisSelecionados.includes(
                item.Material
              )
          );

      }


      return dados;

    }, [
      data,
      mesesSelecionados,
      mesesDoPeriodo,
      coresSelecionadas,
      materiaisSelecionados
    ]);


  /* =======================================================
     KPI
  ======================================================= */

  const totalConsumo =
    useMemo(
      () =>
        dadosFiltrados.reduce(
          (soma, item) =>
            soma +
            numero(
              item.Consumo
            ),
          0
        ),
      [dadosFiltrados]
    );


  const totalCores =
    useMemo(
      () =>
        new Set(
          dadosFiltrados.map(
            x => x.Cor
          )
        ).size,
      [dadosFiltrados]
    );


  const totalMateriais =
    useMemo(
      () =>
        new Set(
          dadosFiltrados.map(
            x => x.Material
          )
        ).size,
      [dadosFiltrados]
    );


  const mediaMensal =
    useMemo(() => {

      const quantidade =
        new Set(
          dadosFiltrados.map(
            x => x.Mes
          )
        ).size;

      return quantidade > 0
        ? totalConsumo /
          quantidade
        : 0;

    }, [
      dadosFiltrados,
      totalConsumo
    ]);


  /* =======================================================
     RANKING DE CORES
  ======================================================= */

  const rankingCores =
    useMemo(() => {

      const totais = {};


      dadosFiltrados.forEach(
        item => {

          if (
            !totais[item.Cor]
          ) {

            totais[item.Cor] = 0;

          }

          totais[item.Cor] +=
            numero(
              item.Consumo
            );

        }
      );


      const total =
        Object.values(
          totais
        ).reduce(
          (a, b) =>
            a + b,
          0
        );


      return Object.entries(
        totais
      )

        .map(
          ([cor, consumo]) => ({

            cor,

            consumo,

            percentual:
              total > 0
                ? (
                    consumo /
                    total
                  ) * 100
                : 0

          })
        )

        .filter(
          item =>
            item.percentual > 1
        )

        .sort(
          (a, b) =>
            b.percentual -
            a.percentual
        );

    }, [
      dadosFiltrados
    ]);


  /* =======================================================
     CONSUMO POR MÊS
  ======================================================= */

  const consumoPorMes =
    useMemo(() => {

      const mapa = {};


      dadosFiltrados.forEach(
        item => {

          if (
            !mapa[item.Mes]
          ) {

            mapa[item.Mes] =
              0;

          }

          mapa[item.Mes] +=
            numero(
              item.Consumo
            );

        }
      );


      return Object.entries(
        mapa
      )
        .map(
          ([mes, consumo]) => ({
            mes,
            consumo
          })
        )
        .sort(
          (a, b) =>
            a.mes.localeCompare(
              b.mes
            )
        );

    }, [
      dadosFiltrados
    ]);


  /* =======================================================
     CONSUMO POR MATERIAL
  ======================================================= */

  const consumoPorMaterial =
    useMemo(() => {

      const mapa = {};


      dadosFiltrados.forEach(
        item => {

          const material = normalizarMaterial(item.Material);

          if (
            !mapa[material]
          ) {

            mapa[material] =
              0;

          }

          mapa[material] +=
            numero(
              item.Consumo
            );

        }
      );


      return Object.entries(
        mapa
      )
        .map(
          ([material, consumo]) => ({
            material,
            consumo
          })
        )
        .sort(
          (a, b) =>
            b.consumo -
            a.consumo
        )
        .slice(0, 15);

    }, [
      dadosFiltrados
    ]);


  /* =======================================================
     HISTÓRICO POR BASE + COR
  ======================================================= */

  const historicoBaseCor =
    useMemo(() => {

      const mapa = {};


      data.forEach(
        item => {

          const material = normalizarMaterial(item.Material);
          const cor = normalizarTexto(item.Cor);

          const chave =
            `${material}|||${cor}`;


          if (
            !mapa[chave]
          ) {

            mapa[chave] = {

              Material:
                material,

              Cor:
                cor,

              Consumo: 0

            };

          }


          mapa[chave].Consumo +=
            numero(
              item.Consumo
            );

        }
      );


      const total =
        Object.values(
          mapa
        ).reduce(
          (soma, item) =>
            soma +
            item.Consumo,
          0
        );


      return Object.values(
        mapa
      )
        .map(
          item => ({

            ...item,

            Percentual:
              total > 0
                ? (
                    item.Consumo /
                    total
                  ) * 100
                : 0

          })
        )
        .sort(
          (a, b) =>
            b.Consumo -
            a.Consumo
        );

    }, [
      data
    ]);


  /* =======================================================
     PROJEÇÃO 2027 FILTRADA PELO MÊS DA ANÁLISE
  ======================================================= */

  const projecaoVendasFiltrada =
    useMemo(() => {

      if (!mesesSelecionados.length) {
        return projecaoVendas;
      }

      const meses2027 =
        mesesSelecionados.map(mes => {
          const texto = String(mes);
          if (/^\d{4}-\d{2}$/.test(texto)) {
            return `2027-${texto.slice(5, 7)}`;
          }
          return texto;
        });

      return projecaoVendas.filter(item =>
        meses2027.includes(item.Mes)
      );

    }, [
      projecaoVendas,
      mesesSelecionados
    ]);


  /* =======================================================
     PROJEÇÃO DE BASES — SOMENTE METROS
  ======================================================= */

  const projecaoBases =
    useMemo(() => {

      if (
        !projecaoVendasFiltrada.length ||
        !historicoBaseCor.length
      ) {
        return [];
      }

      const totalHistorico =
        historicoBaseCor.reduce(
          (soma, item) => soma + item.Consumo,
          0
        );

      if (totalHistorico <= 0) {
        return [];
      }

      const resultado = [];

      projecaoVendasFiltrada.forEach(venda => {

        historicoBaseCor.forEach(item => {

          const percentual =
            item.Consumo / totalHistorico;

          const metrosProjetados =
            venda.Nacional * percentual;

          resultado.push({
            Mes: venda.Mes,
            Material: item.Material,
            Cor: item.Cor,
            Percentual: percentual * 100,
            MetrosProjetados: metrosProjetados
          });

        });

      });

      return resultado;

    }, [
      projecaoVendasFiltrada,
      historicoBaseCor
    ]);


  /* =======================================================
     PROJEÇÃO CONSOLIDADA POR MATERIAL + COR
  ======================================================= */

  const projecaoConsolidada =
    useMemo(() => {

      const mapa = {};

      projecaoBases.forEach(item => {

        const chave =
          `${item.Material}|||${item.Cor}`;

        if (!mapa[chave]) {
          mapa[chave] = {
            Material: item.Material,
            Cor: item.Cor,
            Percentual: item.Percentual,
            MetrosProjetados: 0
          };
        }

        mapa[chave].MetrosProjetados +=
          item.MetrosProjetados;

      });

      return Object.values(mapa).sort(
        (a, b) =>
          b.MetrosProjetados - a.MetrosProjetados
      );

    }, [projecaoBases]);


  const projecaoConsolidadaFiltrada =
    useMemo(() => {

      if (!materiaisProjecaoSelecionados.length) {
        return projecaoConsolidada;
      }

      return projecaoConsolidada.filter(item =>
        materiaisProjecaoSelecionados.includes(item.Material)
      );

    }, [
      projecaoConsolidada,
      materiaisProjecaoSelecionados
    ]);


  /* =======================================================
     PROJEÇÃO POR COR
  ======================================================= */

  const projecaoPorCor =
    useMemo(() => {

      const mapa = {};


      projecaoBases.forEach(
        item => {

          if (
            !mapa[item.Cor]
          ) {

            mapa[item.Cor] = {

              Cor:
                item.Cor,

              Metros: 0

            };

          }


          mapa[item.Cor].Metros +=
            item.MetrosProjetados;

        }
      );


      const total = Object.values(mapa).reduce(
        (soma, item) => soma + item.Metros,
        0
      );

      return Object.values(mapa)
        .map(item => ({
          ...item,
          Percentual: total > 0
            ? (item.Metros / total) * 100
            : 0
        }))
        .sort(
          (a, b) =>
            b.Metros -
            a.Metros
        )
        .slice(0, 15);

    }, [
      projecaoBases
    ]);


  /* =======================================================
     TOTAL PROJETADO
  ======================================================= */

  const totalNacional2027 =
    useMemo(
      () =>
        projecaoVendasFiltrada.reduce(
          (soma, item) => soma + item.Nacional,
          0
        ),
      [projecaoVendasFiltrada]
    );


  const totalMetrosProjetados =
    useMemo(
      () =>
        projecaoBases.reduce(
          (soma, item) =>
            soma +
            item.MetrosProjetados,
          0
        ),
      [projecaoBases]
    );


  /* =======================================================
     RESET
  ======================================================= */

  function limparFiltros() {

    setCoresSelecionadas([]);

    setMateriaisSelecionados([]);

    setMateriaisProjecaoSelecionados([]);

    setMesesSelecionados([]);

    setPeriodo(12);

  }


  /* =======================================================
     LOADING
  ======================================================= */

  if (loading) {

    return (

      <div className="loading-screen">

        <div className="loading-box">
          Carregando dashboard...
        </div>

      </div>

    );

  }


  /* =======================================================
     ERRO
  ======================================================= */

  if (erro) {

    return (

      <div className="error-screen">

        <h2>
          Erro ao carregar os dados
        </h2>

        <p>
          {erro}
        </p>

        <code>
          public/data/Consumo.xlsx
        </code>

      </div>

    );

  }


  /* =======================================================
     DASHBOARD
  ======================================================= */

  return (

    <div className="dashboard">


      {/* =================================================
          HEADER
      ================================================= */}

      <header className="dashboard-header">

        <div>

          <h1>
            Análise de Consumo
          </h1>

          <p>
            Consumo por cor, material e projeção de bases
          </p>

        </div>


        <div className="header-actions">

          <input
            ref={arquivoProjecaoRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={importarPlanilhaProjecao}
            style={{ display: 'none' }}
          />

          <button
            type="button"
            className="import-button"
            onClick={() => arquivoProjecaoRef.current?.click()}
            disabled={atualizandoProjecao}
            title="Atualiza somente a projeção 2027"
          >
            <Upload size={14} />
            {atualizandoProjecao ? 'Atualizando projeção...' : 'Importar projeção 2027'}
          </button>

          <div className="header-period">

            {mesesSelecionados.length > 0
              ? `${mesesSelecionados.length} meses selecionados`
              : `Últimos ${periodo} meses`
            }

          </div>

        </div>

      </header>


      {/* =================================================
          FILTROS
      ================================================= */}

      <section className="filters-card">

        <div className="filters-title">

          <div>

            <Filter size={18} />

            <span>
              Filtros
            </span>

          </div>


          <button
            type="button"
            className="reset-button"
            onClick={
              limparFiltros
            }
          >

            <RotateCcw size={15} />

            Limpar filtros

          </button>

        </div>


        <div className="filters-grid">


          <MultiSelect
            label="Cor"
            icon={Palette}
            options={cores}
            selected={
              coresSelecionadas
            }
            setSelected={
              setCoresSelecionadas
            }
            placeholder="Todas as cores"
          />


          <MultiSelect
            label="Material"
            icon={Layers}
            options={materiais}
            selected={
              materiaisSelecionados
            }
            setSelected={
              setMateriaisSelecionados
            }
            placeholder="Todos os materiais"
          />


          <MultiSelect
            label="Mês"
            icon={TrendingUp}
            options={meses}
            selected={
              mesesSelecionados
            }
            setSelected={
              setMesesSelecionados
            }
            placeholder="Todos os meses"
          />


          <div className="filter-group">

            <label>
              Período
            </label>

            <select
              value={periodo}
              onChange={event =>
                setPeriodo(
                  Number(
                    event.target.value
                  )
                )
              }
              disabled={
                mesesSelecionados.length > 0
              }
            >

              <option value={1}>
                1 mês
              </option>

              <option value={3}>
                3 meses
              </option>

              <option value={6}>
                6 meses
              </option>

              <option value={12}>
                12 meses
              </option>

            </select>

          </div>

        </div>


        {mesesSelecionados.length > 0 && (

          <div className="filter-info">

            Seleção manual de meses ativa.
            O período automático está desabilitado.

          </div>

        )}

        <div className="data-source-info">
          <FileSpreadsheet size={12} />
          <span>{origemDados}</span>
        </div>

      </section>


      {/* =================================================
          KPIs
      ================================================= */}

      <section className="kpi-grid">


        <div className="kpi-card">

          <div className="kpi-icon">
            <Package size={20} />
          </div>

          <div>

            <span>
              Consumo total
            </span>

            <strong>
              {formatarNumero(
                totalConsumo,
                1
              )}
            </strong>

          </div>

        </div>


        <div className="kpi-card">

          <div className="kpi-icon">
            <Palette size={20} />
          </div>

          <div>

            <span>
              Cores
            </span>

            <strong>
              {totalCores}
            </strong>

          </div>

        </div>


        <div className="kpi-card">

          <div className="kpi-icon">
            <Layers size={20} />
          </div>

          <div>

            <span>
              Materiais
            </span>

            <strong>
              {totalMateriais}
            </strong>

          </div>

        </div>


        <div className="kpi-card">

          <div className="kpi-icon">
            <TrendingUp size={20} />
          </div>

          <div>

            <span>
              Média mensal
            </span>

            <strong>
              {formatarNumero(
                mediaMensal,
                1
              )}
            </strong>

          </div>

        </div>

      </section>


      {/* =================================================
          RANKING DE CORES
      ================================================= */}

      <section className="charts-grid">


        <div className="chart-card chart-large">

          <div className="chart-header">

            <div>

              <h2>
                Ranking de cores
              </h2>

              <p>
                Participação percentual no período
              </p>

            </div>

          </div>


          <div className="chart-container color-ranking-chart">

            {rankingCores.length === 0 ? (

              <div className="empty-state">
                Nenhum dado encontrado.
              </div>

            ) : (

              <ResponsiveContainer
                width="100%"
                height={Math.max(
                  420,
                  rankingCores.length *
                    27 +
                    70
                )}
              >

                <BarChart
                  data={
                    rankingCores
                  }
                  layout="vertical"
                  margin={{
                    top: 10,
                    right: 70,
                    left: 15,
                    bottom: 25
                  }}
                  barCategoryGap="22%"
                >

                  <CartesianGrid
                    strokeDasharray="3 3"
                    horizontal={false}
                  />

                  <XAxis
                    type="number"
                    domain={[0, 100]}
                    tickCount={5}
                    tickFormatter={
                      value =>
                        `${value}%`
                    }
                  />

                  <YAxis
                    type="category"
                    dataKey="cor"
                    width={105}
                    interval={0}
                  />

                  <Tooltip
                    cursor={false}
                    formatter={
                      value => [
                        formatarPercentual(
                          value
                        ),
                        'Participação'
                      ]
                    }
                  />

                  <Bar
                    dataKey="percentual"
                    radius={[
                      0,
                      5,
                      5,
                      0
                    ]}
                    barSize={15}
                    activeBar={false}
                  >

                    <LabelList
                      dataKey="percentual"
                      position="right"
                      formatter={
                        value =>
                          formatarPercentual(
                            value
                          )
                      }
                      style={{
                        fontSize: 12,
                        fontWeight: 700,
                        fill: '#1f2937'
                      }}
                    />

                  </Bar>

                </BarChart>

              </ResponsiveContainer>

            )}

          </div>

        </div>
        <div className="chart-card">

          <div className="chart-header">

            <div>

              <h2>
                Projeção por cor — 2027
              </h2>

              <p>
                Distribuição da necessidade projetada em metros
              </p>

            </div>

            <Target size={18} />

          </div>


          <div className="projection-kpis">

            <div>

              <span>
                Necessidade 2027 (m)
              </span>

              <strong>
                {formatarNumero(
                  totalNacional2027,
                  0
                )}
              </strong>

            </div>


            <div>

              <span>
                Meses projetados
              </span>

              <strong>
                {projecaoVendasFiltrada.length}
              </strong>

            </div>

          </div>


          {erroProjecao ? (

            <div className="filter-info">

              {erroProjecao}

            </div>

          ) : (

            <div className="projection-table">

              <div className="projection-row projection-head">

                <span>
                  Cor
                </span>

                <span>
                  Necessidade 2027
                </span>

              </div>


              {projecaoPorCor.map(
                item => (

                  <div
                    className="projection-row"
                    key={item.Cor}
                  >

                    <span>
                      {item.Cor}
                    </span>

                    <strong className="projection-value-with-percent">
                      <span>
                        {formatarNumero(
                          item.Metros,
                          0
                        )} m
                      </span>
                      <small>
                        {formatarPercentual(item.Percentual)}
                      </small>
                    </strong>

                  </div>

                )
              )}

            </div>

          )}

        </div>
        <div className="chart-card">

          <div className="chart-header">

            <div>

              <h2>
                Ranking de materiais
              </h2>

              <p>
                Maiores consumos no período
              </p>

            </div>

          </div>


          <div className="chart-container material-chart">

            <ResponsiveContainer
              width="100%"
              height={520}
            >

              <BarChart
                data={
                  consumoPorMaterial
                }
                layout="vertical"
                margin={{
                  top: 8,
                  right: 45,
                  left: 5,
                  bottom: 8
                }}
                barCategoryGap="18%"
              >

                <CartesianGrid
                  strokeDasharray="3 3"
                  horizontal={false}
                />

                <XAxis
                  type="number"
                />

                <YAxis
                  type="category"
                  dataKey="material"
                  width={210}
                  interval={0}
                  tick={{
                    fontSize: 9
                  }}
                />

                <Tooltip
                  formatter={
                    value =>
                      formatarNumero(
                        value,
                        1
                      )
                  }
                />

                <Bar
                  dataKey="consumo"
                  name="Consumo"
                  radius={[
                    0,
                    5,
                    5,
                    0
                  ]}
                  barSize={16}
                >

                  <LabelList
                    dataKey="consumo"
                    position="right"
                    formatter={
                      value =>
                        formatarNumero(
                          value,
                          0
                        )
                    }
                  />

                </Bar>

              </BarChart>

            </ResponsiveContainer>

          </div>

        </div>
        <div className="chart-card">

          <div className="chart-header">

            <div>

              <h2>
                Consumo por mês
              </h2>

              <p>
                Evolução no período
              </p>

            </div>

          </div>


          <div className="chart-container">

            <ResponsiveContainer
              width="100%"
              height={320}
            >

              <LineChart
                data={
                  consumoPorMes
                }
              >

                <CartesianGrid
                  strokeDasharray="3 3"
                />

                <XAxis
                  dataKey="mes"
                />

                <YAxis />

                <Tooltip
                  formatter={
                    value =>
                      formatarNumero(
                        value,
                        1
                      )
                  }
                />

                <Line
                  type="monotone"
                  dataKey="consumo"
                  strokeWidth={3}
                  dot
                  name="Consumo"
                />

              </LineChart>

            </ResponsiveContainer>

          </div>

        </div>
</section>


      {/* =================================================
          TABELA FINAL
      ================================================= */}

      <section className="chart-card projection-section">

        <div className="chart-header">

          <div>

            <h2>
              Necessidade projetada por base e cor
            </h2>

            <p>
              Necessidade projetada em metros para 2027
            </p>

          </div>

          <BarChart3 size={18} />

        </div>

        <div className="projection-filter-row">

          <MultiSelect
            label="Material"
            icon={Layers}
            options={materiais}
            selected={materiaisProjecaoSelecionados}
            setSelected={setMateriaisProjecaoSelecionados}
            placeholder="Todos os materiais"
          />

        </div>


        <div className="projection-table wide">

          <div className="projection-row projection-head">

            <span>
              Tipo de base
            </span>

            <span>
              Cor
            </span>

            <span>
              %
            </span>

            <span>
              Necessidade 2027
            </span>

          </div>


          {projecaoConsolidadaFiltrada.map(
            item => (

              <div
                className="projection-row"
                key={`${item.Material}-${item.Cor}`}
              >

                <span>
                  {item.Material}
                </span>

                <span>
                  {item.Cor}
                </span>

                <strong>
                  {formatarPercentual(
                    item.Percentual
                  )}
                </strong>

                <strong>
                  {formatarNumero(
                    item.MetrosProjetados,
                    0
                  )} m
                </strong>

              </div>

            )
          )}

        </div>

      </section>




    </div>

  );

}