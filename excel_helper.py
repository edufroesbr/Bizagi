import os
import glob
import logging
import time
import win32com.client

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Constants (Paths provided by user)
MASTER_LIST_PATH = r"C:\Users\CLIENTE\OneDrive\Documentos\ONS\BIZAGI\Contratos Rescindidos - REN 1125.xlsx"
AVD_FOLDER_PATH = r"C:\Users\CLIENTE\OneDrive\Documentos\ONS\BIZAGI\AVD Complementares\AVDs complementares"

def get_avd_file_path(reference_number):
    """
    Finds the path to the AVD excel file containing the reference_number in its name.
    """
    if not os.path.exists(AVD_FOLDER_PATH):
        logger.error(f"Pasta AVD não encontrada: {AVD_FOLDER_PATH}")
        return None
        
    search_pattern = os.path.join(AVD_FOLDER_PATH, f"*{reference_number}*.xlsx")
    files = glob.glob(search_pattern)
    
    if files:
        return files[0]
    return None

def find_reference_number_visual(contract_code):
    """
    Control Excel to visually find the contract code and return the ID.
    1. Open Master List Visible
    2. Filter Column C (Contrato)
    3. Highlight Column A (ID)
    4. Return ID
    """
    try:
        logger.info("--- [EXCEL VISUAL] Iniciando automação... ---")
        
        # Ensure absolute path (COM requires full paths)
        abs_path = os.path.abspath(MASTER_LIST_PATH)
        logger.info(f"--- [EXCEL VISUAL] Caminho absoluto: {abs_path}")
        
        if not os.path.exists(abs_path):
            logger.error(f"FATAL: Arquivo não encontrado no disco: {abs_path}")
            return None

        # Connect to Excel
        logger.info("--- [EXCEL VISUAL] Conectando ao Excel (Dispatch)... ---")
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            logger.info("--- [EXCEL VISUAL] Excel conectado. Tornando visível... ---")
            excel.Visible = True
            excel.DisplayAlerts = False 
        except Exception as e:
            logger.error(f"FATAL: Falha ao iniciar Excel: {e}")
            return None
        
        # Open Workbook
        logger.info(f"--- [EXCEL VISUAL] Abrindo pasta de trabalho: {os.path.basename(abs_path)}... ---")
        try:
            wb = excel.Workbooks.Open(abs_path)
            logger.info("--- [EXCEL VISUAL] Pasta de trabalho aberta! ---")
        except Exception as e:
             logger.error(f"FATAL: Falha ao abrir arquivo (pode estar em uso/travado): {e}")
             return None

        ws = wb.Sheets(1) # Assume first sheet
        
        # Ensure Clean State
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            
        logger.info(f"--- [EXCEL VISUAL] Filtrando contrato: {contract_code} ---")
        ws.Range("A1").AutoFilter(Field=3, Criteria1=contract_code)
        
        # Find visible cell in Column A
        try:
            visible_range = ws.UsedRange.SpecialCells(12)
            found_id = None
            
            for area in visible_range.Areas:
                for row in area.Rows:
                    if row.Row == 1: continue # Skip header
                    
                    cell_id = ws.Cells(row.Row, 1) # Col A
                    cell_contract = ws.Cells(row.Row, 3) # Col C
                    
                    val = cell_id.Value
                    found_id = str(int(val)) if isinstance(val, float) else str(val)
                    
                    # HIGHLIGHT
                    logger.info(f"--- [EXCEL VISUAL] Destacando linha {row.Row} (ID: {found_id}) ---")
                    cell_id.Interior.Color = 65535 # Yellow
                    cell_contract.Interior.Color = 65535 
                    
                    excel.ActiveWindow.ScrollRow = row.Row
                    break 
                if found_id: break
                
            if found_id:
                return found_id
            else:
                logger.warning("--- [EXCEL VISUAL] Nenhuma linha visível após filtro. ---")
                return None
                
        except Exception as e:
            logger.error(f"Erro ao processar células visíveis: {e}")
            return None
            
    except Exception as e:
        logger.error(f"Erro geral na automação Excel: {e}")
        return None

def validate_debt_amount_visual(reference_number, target_cnpj, bizagi_amount):
    """
    Opens AVD file visually, filters by CNPJ, highlights valid rows.
    """
    try:
        avd_path = get_avd_file_path(reference_number)
        if not avd_path:
            return False, 0.0, f"Arquivo AVD não encontrado para ID {reference_number}"
        
        abs_path = os.path.abspath(avd_path)
        logger.info(f"--- [AVD VISUAL] Abrindo arquivo: {abs_path} ---")
        
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True
        
        try:
            wb = excel.Workbooks.Open(abs_path)
        except Exception as e:
             return False, 0.0, f"Erro ao abrir AVD (COM): {e}"

        ws = wb.Sheets(1)
        
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
        if ws.AutoFilterMode:
            ws.AutoFilterMode = False
            
        # USER REQUEST: Search using Standard Format "XX.XXX.XXX/XXXX-XX"
        raw_cnpj = ''.join(filter(str.isdigit, target_cnpj))
        
        def format_cnpj(cnpj):
            if len(cnpj) != 14: return cnpj
            return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
            
        formatted_cnpj = format_cnpj(raw_cnpj)
        
        # 1. FIND THE HEADER ROW
        # The user screenshot shows "CNPJ" is in Row 20, Col C.
        # We search for the cell containing "CNPJ" explicitly to handle variable layouts.
        logger.info("--- [AVD VISUAL] Localizando cabeçalho 'CNPJ'... ---")
        
        header_cell = ws.UsedRange.Find("CNPJ")
        
        if not header_cell:
            logger.error("--- [AVD VISUAL] Cabeçalho 'CNPJ' não encontrado! Impossível filtrar. ---")
            return False, 0.0, "Cabeçalho CNPJ não encontrado."
            
        header_row = header_cell.Row
        cnpj_col_index = header_cell.Column # Should be 3 (C)
        
        logger.info(f"--- [AVD VISUAL] Cabeçalho encontrado na Linha {header_row}, Coluna {cnpj_col_index} ---")
        
        # 2. DEFINING THE TABLE RANGE
        # We start the filter range from the beginning of that row (Col A) to the end.
        # Ideally, use CurrentRegion around the header, or just the row.
        # Be careful with merged cells above. We want to filter *from* this row downwards.
        
        # Determine the table range starting from the header row
        # We assume the table covers at least the columns we notice.
        # Safe bet: Apply filter to the specific header row.
        
        # ws.Rows(header_row).AutoFilter() 
        
        # 3. APPLY FILTER
        # Field is relative to the first column of the range.
        # If we filter the whole row, Field is the column index (e.g. 3 for C).
        
        criteria = formatted_cnpj
        
        logger.info(f"--- [AVD VISUAL] Aplicando Filtro na Linha {header_row}, Field {cnpj_col_index} com {criteria} ---")
        
        # Attempt fallback if initial search fails? 
        # Actually AutoFilter needs exact match usually, or *cards.
        # The user screenshot shows exact format.
        
        try:
            ws.Rows(header_row).AutoFilter(Field=cnpj_col_index, Criteria1=criteria)
        except Exception as e:
            # Fallback: Try with current region if Rows() fails
            logger.warning(f"AutoFilter na linha falhou: {e}. Tentando CurrentRegion...")
            header_cell.CurrentRegion.AutoFilter(Field=cnpj_col_index, Criteria1=criteria)
        
        # 4. HIGHLIGHT & CALCULATE SUM (Column K - "Total")
        
        total_col_index = 11 # Default K
        
        # Search for "Total" strictly in the Header Row to avoid "Total de Pagamentos" in metadata
        header_total = ws.Rows(header_row).Find("Total")
        
        if header_total:
            total_col_index = header_total.Column
            logger.info(f"--- [AVD VISUAL] Coluna 'Total' encontrada na Linha {header_row}, Coluna {total_col_index} ---")
        else:
            logger.info("--- [AVD VISUAL] Coluna 'Total' não encontrada no cabeçalho. Tentando busca exata 'Total'... ---")
            # Try exact match loop if Find is partial? Or just stick to default.
            logger.warning("--- [AVD VISUAL] Usando padrão Coluna K (11). ---")
        
        visible_range = ws.UsedRange.SpecialCells(12)
        row_count = 0
        total_debt_sum = 0.0
        
        for area in visible_range.Areas:
            for row in area.Rows:
                if row.Row <= header_row: continue 
                
                # Highlight Row
                ws.Rows(row.Row).Interior.Color = 65535 
                row_count += 1
                
                # Sum Value from Total Column
                try:
                    cell_val = ws.Cells(row.Row, total_col_index).Value
                    if isinstance(cell_val, (int, float)):
                        total_debt_sum += float(cell_val)
                except Exception as e:
                    logger.warning(f"Erro ao ler valor da linha {row.Row}: {e}")

        logger.info(f"--- [AVD VISUAL] Soma Total Encontrada: {total_debt_sum} ---")
        
        # Compare with Bizagi Amount
        # Clean Bizagi Amount (R$ 4.846,53 -> 4846.53)
        bizagi_float = 0.0
        try:
            clean_bizagi = str(bizagi_amount).replace('R$', '').replace('.', '').replace(',', '.').strip()
            bizagi_float = float(clean_bizagi)
        except:
            logger.warning(f"Não foi possível converter valor do Bizagi: {bizagi_amount}")
        
        # Tolerance check
        diff = abs(total_debt_sum - bizagi_float)
        is_valid = diff < 1.0 # Tolerance of 1 real
        
        msg = f"Soma Planilha: {total_debt_sum:.2f} | Bizagi: {bizagi_float:.2f} | Diferença: {diff:.2f}"
        
        if is_valid:
             return True, total_debt_sum, f"SUCESSO: Valores conferem! {msg}"
        else:
             return False, total_debt_sum, f"DIVERGÊNCIA: {msg}"

    except Exception as e:
        logger.error(f"Erro na visualização AVD: {e}")
        return False, 0.0, str(e)
        row_count = 0
        
        for area in visible_range.Areas:
            for row in area.Rows:
                if row.Row == 1: continue 
                
                # Highlight
                ws.Rows(row.Row).Interior.Color = 65535 
                row_count += 1
        
        if row_count > 0:
             return True, 0.0, f"Visualização concluída. {row_count} linhas destacadas."
        else:
             return True, 0.0, "CNPJ encontrado mas filtro ocultou linhas? (Incomum)."

    except Exception as e:
        logger.error(f"Erro na visualização AVD: {e}")
        return False, 0.0, str(e)
