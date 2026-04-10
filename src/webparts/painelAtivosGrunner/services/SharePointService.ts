import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/site-users/web"; 
import "@pnp/sp/profiles"; 

const COLUNA_ID_PATRIMONIO = "field_4"; 
const COLUNA_PATRIMONIO_FINANCEIRO = "field_5"; 
const COLUNA_IMEI = "field_9";
const COLUNA_ESPECIFICACOES = "field_10";
const COLUNA_RESPONSAVEL_AD = "Responsavel_AD"; 

export class SharePointService {
  private _sp: SPFI;

  constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  public async getProximoIdSequencial(tipoAtivo: string): Promise<string> {
    try {
      const letraBase = tipoAtivo === "Celular / Smartphone" ? "C" : "N";

      const itens = await this._sp.web.lists.getByTitle("Ativos de TI")
        .items.select(COLUNA_ID_PATRIMONIO)
        .orderBy("Created", false)
        .top(50)();

      if (!itens || itens.length === 0) return `${letraBase}0001`;

      let ultimoCodigoValido = "";
      for (const item of itens) {
        const valorCodigo = item[COLUNA_ID_PATRIMONIO];
        if (valorCodigo && typeof valorCodigo === "string" && valorCodigo.startsWith(letraBase)) {
          ultimoCodigoValido = valorCodigo;
          break; 
        }
      }

      if (!ultimoCodigoValido) return `${letraBase}0001`;
      
      const numeroApenas = parseInt(ultimoCodigoValido.replace(/\D/g, ""), 10);
      const proximoNumero = isNaN(numeroApenas) ? 1 : numeroApenas + 1;

      return `${letraBase}${proximoNumero.toString().padStart(4, '0')}`;
    } catch (error) {
      console.error("Erro ao gerar ID Sequencial:", error);
      throw error;
    }
  }

  public async salvarNovoAtivo(itemDoCarrinho: any, nomeResponsavel: string, departamento: string, emailResponsavel: string): Promise<{ codigo: string }> {
    try {
      const novoCodigo = await this.getProximoIdSequencial(itemDoCarrinho.tipo);
      const primeiraLetra = novoCodigo.charAt(0).toUpperCase();
      
      const payload: any = {
        Title: nomeResponsavel,         
        field_1: departamento,
        field_2: itemDoCarrinho.tipo,        
        field_3: primeiraLetra,         
        field_6: itemDoCarrinho.fabricante,  
        field_7: itemDoCarrinho.modelo,      
        field_8: itemDoCarrinho.serie,       
        field_11: itemDoCarrinho.observacoes || "Sem observações adicionais" 
      };

      payload[COLUNA_ID_PATRIMONIO] = novoCodigo;
      payload[COLUNA_PATRIMONIO_FINANCEIRO] = itemDoCarrinho.patrimonioFin;
      payload[COLUNA_IMEI] = itemDoCarrinho.imei;
      payload[COLUNA_ESPECIFICACOES] = itemDoCarrinho.especificacoes;

      if (emailResponsavel) {
        try {
          const user = await this._sp.web.ensureUser(emailResponsavel);
          payload[`${COLUNA_RESPONSAVEL_AD}Id`] = user.Id; 
        } catch (err) {
          console.warn("Não foi possível validar o usuário no AD:", err);
        }
      }

      await this._sp.web.lists.getByTitle("Ativos de TI").items.add(payload);

      return { codigo: novoCodigo };
    } catch (error) {
      console.error("Erro ao salvar no SharePoint:", error);
      throw error;
    }
  }

  public async atualizarAtivo(id: number, dados: any, emailResponsavel: string): Promise<void> {
    try {
      const payload: any = {
        Title: dados.nome,         
        field_1: dados.departamento,
        field_2: dados.tipo,             
        field_6: dados.fabricante,  
        field_7: dados.modelo,      
        field_8: dados.serie,       
        field_11: dados.observacao || "Sem observações"
      };

      payload[COLUNA_PATRIMONIO_FINANCEIRO] = dados.patrimonioFin;
      payload[COLUNA_IMEI] = dados.imei;
      payload[COLUNA_ESPECIFICACOES] = dados.especificacao;

      if (emailResponsavel) {
        try {
          const user = await this._sp.web.ensureUser(emailResponsavel);
          payload[`${COLUNA_RESPONSAVEL_AD}Id`] = user.Id; 
        } catch (err) {
          console.warn("Não foi possível validar o usuário no AD na edição:", err);
        }
      } else {
        payload[`${COLUNA_RESPONSAVEL_AD}Id`] = null; 
      }

      await this._sp.web.lists.getByTitle("Ativos de TI").items.getById(id).update(payload);
    } catch (error) {
      console.error("Erro ao atualizar no SharePoint:", error);
      throw error;
    }
  }

  public async transferirAtivo(id: number, nomeResponsavel: string, departamento: string, emailResponsavel: string, observacoes: string): Promise<void> {
    try {
      const payload: any = {
        Title: nomeResponsavel,
        field_1: departamento,
        field_11: observacoes || "Transferido em lote"
      };

      if (emailResponsavel) {
        try {
          const user = await this._sp.web.ensureUser(emailResponsavel);
          payload[`${COLUNA_RESPONSAVEL_AD}Id`] = user.Id; 
        } catch (err) {
          console.warn("Não foi possível validar o usuário no AD:", err);
        }
      } else {
        payload[`${COLUNA_RESPONSAVEL_AD}Id`] = null; 
      }

      await this._sp.web.lists.getByTitle("Ativos de TI").items.getById(id).update(payload);
    } catch (error) {
      console.error("Erro ao transferir no SharePoint:", error);
      throw error;
    }
  }

  public async getTemplateTermo(): Promise<ArrayBuffer> {
    try {
      return await this._sp.web
        .getFileByServerRelativePath("/sites/IntranetGrunner/Modelos_TI/Molde_Termo_Grunner.docx")
        .getBuffer();
    } catch (error) {
      console.error("Erro ao buscar template do Word:", error);
      throw error;
    }
  }

  public async getTodosAtivos(): Promise<any[]> {
    try {
      const itens = await this._sp.web.lists.getByTitle("Ativos de TI")
        .items
        .select("Id", "Title", "field_1", "field_2", "field_6", "field_7", "field_8", "field_11", "Created", `${COLUNA_RESPONSAVEL_AD}/Title`, `${COLUNA_RESPONSAVEL_AD}/EMail`, COLUNA_ID_PATRIMONIO, COLUNA_PATRIMONIO_FINANCEIRO, COLUNA_IMEI, COLUNA_ESPECIFICACOES)
        .expand(COLUNA_RESPONSAVEL_AD)
        .orderBy("Created", false) 
        .top(1000)(); 

      return itens.map((item: any) => ({
        id: item.Id,
        responsavel: (item[COLUNA_RESPONSAVEL_AD] && item[COLUNA_RESPONSAVEL_AD].Title) ? item[COLUNA_RESPONSAVEL_AD].Title : item.Title || "Sem Responsável",
        emailResponsavel: (item[COLUNA_RESPONSAVEL_AD] && item[COLUNA_RESPONSAVEL_AD].EMail) ? item[COLUNA_RESPONSAVEL_AD].EMail : "",
        departamento: item.field_1 || "",
        tipo: item.field_2 || "",
        patrimonio: item[COLUNA_ID_PATRIMONIO] || "-",
        patrimonioFin: item[COLUNA_PATRIMONIO_FINANCEIRO] || "-",
        fabricante: item.field_6 || "",
        modelo: item.field_7 || "",
        serie: item.field_8 || item[COLUNA_IMEI] || "-", 
        especificacoes: item[COLUNA_ESPECIFICACOES] || "",
        observacoes: item.field_11 || "",
        dataCriacao: new Date(item.Created).toLocaleDateString('pt-BR')
      }));
    } catch (error) {
      console.error("Erro ao buscar a lista de ativos:", error);
      throw error;
    }
  }

  public async buscarUsuariosAD(termo: string): Promise<any[]> {
    if (!termo || termo.length < 3) return [];
    try {
      const usuarios = await this._sp.web.siteUsers
        .filter(`substringof('${termo}', Title) or substringof('${termo}', Email)`)
        .top(5)();
      
      return usuarios.map((u: any) => ({
        id: u.Id,
        nome: u.Title,
        email: u.Email || u.LoginName
      }));
    } catch (error) {
      console.warn("Erro ao buscar usuários:", error);
      return [];
    }
  }

  public async getDepartamentoUsuario(email: string): Promise<string> {
    try {
      const loginName = `i:0#.f|membership|${email}`;
      const profile: any = await this._sp.profiles.getPropertiesFor(loginName);
      const propriedades = profile.UserProfileProperties || (profile.data && profile.data.UserProfileProperties);

      if (propriedades) {
        const dep = propriedades.find((p: any) => p.Key === "Department");
        return dep && dep.Value ? dep.Value : "";
      }
      return ""; 
    } catch (error) {
      console.warn("Não foi possível buscar o departamento no AD:", error);
      return "";
    }
  }

  public async getHistoricoAtivo(id: number): Promise<any[]> {
    try {
      const versoes = await this._sp.web.lists.getByTitle("Ativos de TI").items.getById(id).versions();
      return versoes.map((v: any) => ({
        versao: v.VersionLabel,
        data: new Date(v.Created).toLocaleString('pt-BR'),
        modificadoPor: v.Editor ? (v.Editor.LookupValue || v.Editor.Email) : "Sistema",
        responsavel: v.Title || "Sem Responsável",
        observacao: v.field_11 || ""
      }));
    } catch (error) {
      console.warn("Não foi possível buscar o histórico de versões. Verifique se o Versionamento está ativo na lista.", error);
      return [];
    }
  }

  // --- NOVA FUNÇÃO: LER LISTA DINÂMICA DE ACESSOS ---
  public async verificarAcessoUsuario(emailLogado: string): Promise<{ isTI: boolean, isVisualizador: boolean }> {
    try {
      // Puxa todos os itens da nossa nova lista de acessos
      const itens = await this._sp.web.lists.getByTitle("Acessos_Painel_Ativos").items.select("Email", "NivelAcesso")();
      
      let isTI = false;
      let isVisualizador = false;

      for (const item of itens) {
        // Validação ignorando letras maiúsculas/minúsculas e espaços
        if (item.Email && item.Email.trim().toLowerCase() === emailLogado.trim().toLowerCase()) {
          if (item.NivelAcesso === "TI") {
            isTI = true;
          } else if (item.NivelAcesso === "Visualizador") {
            isVisualizador = true;
          }
        }
      }
      
      return { isTI, isVisualizador };
    } catch (error) {
      console.warn("Erro ao buscar acessos na lista 'Acessos_Painel_Ativos'. O usuário será tratado como Colaborador comum por segurança.", error);
      return { isTI: false, isVisualizador: false };
    }
  }

  // --- FUNÇÕES DE GERENCIAMENTO DE ACESSOS ---
  public async getTodosAcessos(): Promise<any[]> {
    try {
      const itens = await this._sp.web.lists.getByTitle("Acessos_Painel_Ativos").items.select("Id", "Title", "Email", "NivelAcesso")();
      return itens.map((item: any) => ({
        id: item.Id,
        nome: item.Title || "",
        email: item.Email || "",
        nivel: item.NivelAcesso || "Visualizador"
      }));
    } catch (error) {
      console.error("Erro ao buscar a lista de acessos:", error);
      return [];
    }
  }

  public async adicionarAcesso(nome: string, email: string, nivel: string): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle("Acessos_Painel_Ativos").items.add({
        Title: nome,
        Email: email,
        NivelAcesso: nivel
      });
    } catch (error) {
      console.error("Erro ao salvar novo acesso:", error);
      throw error;
    }
  }

  public async removerAcesso(id: number): Promise<void> {
    try {
      await this._sp.web.lists.getByTitle("Acessos_Painel_Ativos").items.getById(id).delete();
    } catch (error) {
      console.error("Erro ao remover acesso:", error);
      throw error;
    }
  }

}