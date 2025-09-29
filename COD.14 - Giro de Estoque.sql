SELECT
    P.COD_ITEM AS Cod_Item,
    P.DESC_ITEM AS Descricao,
    MAX(KDX.DATA) AS Ultima_compra,
    DATEDIFF(DAY, MAX(KDX.DATA), GETDATE()) AS dias_sem_giro,
    CASE
	WHEN DATEDIFF(DAY, MAX(KDX.DATA), GETDATE()) < 180 THEN 'Produto OK'
        WHEN DATEDIFF(DAY, MAX(KDX.DATA), GETDATE()) >= 180 AND <= 720 THEN 'Baixo Giro'
        WHEN DATEDIFF(DAY, MAX(KDX.DATA), GETDATE()) > 720 THEN 'Produto Obsoleto' 
        ELSE 'N/A'
    END AS Giro_de_Estoque
    CASE
	WHEN QTD < 0 THEN 'Saída'
	WHEN QTD > 0 THEN 'Entrada'
	ELSE 'Sem_Movimentacao'
    END Movimentacao
FROM
    VW_AUDIT_RM_ESTOQUE_PRODUTO AS P
INNER JOIN
    VW_AUDIT_RM_GERENCIAMENTO_ESTOQUE AS KDX ON P.COD_ITEM = KDX.COD_ITEM
WHERE
    KDX.tipo_movimentacao = 'Entrada'
    AND DEPOSITO LIKE '%BTN%'	
GROUP BY
    P.COD_ITEM, P.DESC_ITEM
