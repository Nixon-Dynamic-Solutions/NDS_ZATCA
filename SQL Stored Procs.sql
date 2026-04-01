/* =========================================================
   1. TaxType
   ========================================================= */
IF OBJECT_ID('dbo.TaxType', 'P') IS NOT NULL
    DROP PROCEDURE dbo.TaxType;
GO

CREATE PROCEDURE dbo.TaxType
    @Tax NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;

    IF @Tax = ''
    BEGIN
        SELECT ' -- ' AS Code, ' -- ' AS Name
        UNION ALL
        SELECT U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc
        FROM [@I_ZATCA_TAXCODE];
    END
    ELSE IF @Tax = 'KEXO'
    BEGIN
        SELECT ' -- ' AS Code, ' -- ' AS Name
        UNION ALL
        SELECT U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc
        FROM [@I_ZATCA_TAXCODE]
        WHERE U_I_Tax_Ex_Type_Code = 'E';
    END
    ELSE IF @Tax = 'KOCO' OR @Tax = '02'
    BEGIN
        SELECT ' -- ' AS Code, ' -- ' AS Name
        UNION ALL
        SELECT U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc
        FROM [@I_ZATCA_TAXCODE]
        WHERE U_I_Tax_Ex_Type_Code = 'Z';
    END
END;
GO


/* =========================================================
   2. @EINVOICE_DETAIL
   ========================================================= */
IF OBJECT_ID('dbo.[@EINVOICE_DETAIL]', 'P') IS NOT NULL
    DROP PROCEDURE dbo.[@EINVOICE_DETAIL];
GO

CREATE PROCEDURE dbo.[@EINVOICE_DETAIL]
    @DocEntry NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT *
    FROM
    (
        SELECT
            B.ItemCode,
            B.Dscription,
            B.VatSum,
            B.VatGroup,

            ROUND(
                (B.Price /
                    CASE
                        WHEN B.Rate = 0.0 THEN 1.0
                        ELSE B.Rate
                    END
                ), 2
            ) AS Price,

            B.LineNum,

            CASE
                WHEN A.DocType = 'I' THEN B.Quantity
                WHEN A.DocType = 'S' THEN 1
            END AS Quantity,

            B.U_APIRefNo,
            B.Currency,
            'PCE' AS Unitmsr,

            CASE
                WHEN B.DiscPrcnt > 0
                THEN B.PriceBefDi * (B.DiscPrcnt / 100.0)
                ELSE 0
            END AS Discount,

            B.LineTotal,

            ROUND(B.LineTotal * (C.Rate / 100.0), 2) AS vat,

            B.VatSum AS v,

            ROUND(
                (B.Price /
                    CASE
                        WHEN B.Rate = 0.0 THEN 1.0
                        ELSE B.Rate
                    END
                ), 2
            ) AS BaseAmount,

            B.DiscPrcnt

        FROM INV1 B
        INNER JOIN OINV A
            ON A.DocEntry = B.DocEntry
        INNER JOIN OVTG C
            ON B.VatGroup = C.Code
        WHERE A.DocEntry = @DocEntry
    ) A;
END;
GO


/* =========================================================
   3. @EINVOICE_HEADER
   ========================================================= */
IF OBJECT_ID('dbo.[@EINVOICE_HEADER]', 'P') IS NOT NULL
    DROP PROCEDURE dbo.[@EINVOICE_HEADER];
GO

CREATE PROCEDURE dbo.[@EINVOICE_HEADER]
    @DocEntry NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        A.DocNum,
        A.CardName,
        A.DocDate,

        RIGHT('000000' + CAST(A.CreateTS AS VARCHAR(6)), 6) AS CreateTSRaw,
        SUBSTRING(RIGHT('000000' + CAST(A.CreateTS AS VARCHAR(6)), 6), 1, 2) + ':' +
        SUBSTRING(RIGHT('000000' + CAST(A.CreateTS AS VARCHAR(6)), 6), 3, 2) + ':' +
        SUBSTRING(RIGHT('000000' + CAST(A.CreateTS AS VARCHAR(6)), 6), 5, 2) AS CreateTime,

        'SAR' AS DocCur,

        C.U_District,
        D.StreetB,
        D.CityB,
        D.CountryB,
        D.ZipCodeB,
        D.BuildingB,

        E.TaxPayerRf,
        E.CompnyName,

        'SAR' AS TaxCur,

        A.Comments,

        C.Building AS CBuilding,
        C.City AS CCity,
        C.U_AddNo,
        C.ZipCode AS CZipCode,
        C.State AS CState,
        C.Country AS CCountry,

        B.LicTradNum AS TaxIDNum3,

        CASE
            WHEN B.RegNum = '-' THEN ''
            ELSE B.RegNum
        END AS RegNum,

        C.Street AS CStreet,

        A.NumAtCard,
        A.U_CustRef,
        A.U_PS_SDate,

        A.DiscPrcnt,

        A.U_ZATCA_TaxCode AS TaxReasonCode,
        T.U_I_Tax_Ex_Desc AS TaxReason,

        F.Building,
        F.Street,
        F.City,
        F.ZipCode,
        F.County,
        F.Country,
        F.State,

        '8CC649D2-7802-4BEF-A8B5-CA21C82F70A4' AS UUID,

        G.[Count],

        V.Rate,

        H.LineAmnt,
        H.LineAmnt - A.DiscSum AS Taxexamnt,
        (H.LineAmnt - A.DiscSum) + A.VatSum AS Taxinamnt,

        A.DiscSum,
        A.DiscPrcnt,

        H.LineAmnt - A.DiscSum AS BaseAmount,
        H.LineAmnt - A.DiscSum AS Total1,

        A.VatSum,
        A.DocTotal

    FROM OINV A
    INNER JOIN OCRD B
        ON A.CardCode = B.CardCode
    LEFT JOIN CRD1 C
        ON B.CardCode = C.CardCode
    INNER JOIN INV12 D
        ON A.DocEntry = D.DocEntry
    INNER JOIN
    (
        SELECT COUNT(*) AS [Count], DocEntry
        FROM INV1
        GROUP BY DocEntry
    ) G
        ON A.DocEntry = G.DocEntry
    INNER JOIN
    (
        SELECT
            SUM(LineTotal) AS LineAmnt,
            SUM(
                ROUND(
                    (B.PriceBefDi *
                        CASE
                            WHEN A.DocType = 'I' THEN B.Quantity
                            WHEN A.DocType = 'S' THEN 1
                        END
                    ) /
                    CASE
                        WHEN B.Rate = 0.0 THEN 1.0
                        ELSE B.Rate
                    END
                ,2)
            ) AS Base,
            B.DocEntry
        FROM INV1 B
        LEFT JOIN OINV A
            ON A.DocEntry = B.DocEntry
        GROUP BY B.DocEntry
    ) H
        ON A.DocEntry = H.DocEntry
    INNER JOIN
    (
        SELECT MAX(B.Rate) AS Rate, A.DocEntry
        FROM INV1 A
        INNER JOIN OVTG B
            ON A.VatGroup = B.Code
        GROUP BY A.DocEntry
    ) V
        ON A.DocEntry = V.DocEntry
    LEFT JOIN [@I_ZATCA_TAXCODE] T
        ON A.U_ZATCA_TaxCode = T.U_I_Tax_Ex_Code
    CROSS JOIN OADM E
    CROSS JOIN ADM1 F
    WHERE A.DocEntry = @DocEntry
      AND C.AdresType = 'B';
END;
GO


/* =========================================================
   4. ZATCA Tax Code VALUES
   ========================================================= */
INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('12', '12', 'VATEX-SA-29', 'Financial services mentioned in Article 29 of the VAT Regulations', 'Exempt from Tax', 'E');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('13', '13', 'VATEX-SA-29-7', 'Life insurance services mentioned in Article 29 of the VAT Regulations', 'Exempt from Tax', 'E');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('14', '14', 'VATEX-SA-30', 'Real estate transactions mentioned in Article 30 of the VAT Regulations', 'Exempt from Tax', 'E');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('1', '1', 'VATEX-SA-32', 'Export of goods', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('10', '10', 'VATEX-SA-EDU', 'Private education to citizen', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('11', '11', 'VATEX-SA-HEA', 'Private healthcare to citizen', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('2', '2', 'VATEX-SA-33', 'Export of services', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('3', '3', 'VATEX-SA-34-1', 'Export of services', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('4', '4', 'VATEX-SA-34-2', 'international transport of passengers', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('5', '5', 'VATEX-SA-34-3', 'international passenger transport', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('6', '6', 'VATEX-SA-34-4', 'Supply of a qualifying means of transport', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('7', '7', 'VATEX-SA-34-5', 'Any services relating to Goods or passenger transportation, as defined in article twenty five of these Regulations', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('8', '8', 'VATEX-SA-35', 'Medicines and medical equipment', 'zero rated goods', 'Z');

INSERT INTO [@I_ZATCA_TAXCODE]
(Code, Name, U_I_Tax_Ex_Code, U_I_Tax_Ex_Desc, U_I_Tax_Ex_Type, U_I_Tax_Ex_Type_Code)
VALUES ('9', '9', 'VATEX-SA-36', 'Qualifying metals', 'zero rated goods', 'Z');
GO


/* =========================================================
   5. @ECREDITMEMO_HEADER
   ========================================================= */
IF OBJECT_ID('dbo.[@ECREDITMEMO_HEADER]', 'P') IS NOT NULL
    DROP PROCEDURE dbo.[@ECREDITMEMO_HEADER];
GO

CREATE PROCEDURE dbo.[@ECREDITMEMO_HEADER]
    @DocEntry NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        a.DocNum,
        a.CardName,
        a.DocDate,

        SUBSTRING(RIGHT('000000' + CAST(a.CreateTS AS VARCHAR(6)), 6), 1, 2) + ':' +
        SUBSTRING(RIGHT('000000' + CAST(a.CreateTS AS VARCHAR(6)), 6), 3, 2) + ':' +
        SUBSTRING(RIGHT('000000' + CAST(a.CreateTS AS VARCHAR(6)), 6), 5, 2) AS CreateTime,

        'SAR' AS DocCur,
        c.U_District,
        d.StreetB,
        d.CityB,
        d.CountryB,
        d.ZipCodeB,
        d.BuildingB,
        e.TaxPayerRf,
        e.CompnyName,
        'SAR' AS TaxCur,
        a.Comments,
        c.Building AS CBuilding,
        c.City AS CCity,
        c.U_AddNo,
        c.ZipCode AS CZipCode,
        c.State AS CState,
        c.Country AS CCountry,
        b.LicTradNum AS TaxIDNum3,

        CASE
            WHEN b.RegNum = '-' THEN ''
            ELSE b.RegNum
        END AS RegNum,

        c.Street AS CStreet,
        a.NumAtCard,
        a.U_CustRef,
        a.U_PS_SDate,
        a.DiscPrcnt,

        g.BaseRef AS InvNo,
        a.U_Rem,

        f.Building,
        f.Street,
        f.City,
        f.ZipCode,
        f.County,
        f.Country,
        f.State,

        '8CC649D2-7802-4BEF-A8B5-CA21C82F70A4' AS UUID,

        g.[Count],

        a.U_ZATCA_TaxCode AS TaxReasonCode,
        t.U_I_Tax_Ex_Desc AS TaxReason,

        v.Rate,

        h.LineAmnt,
        h.LineAmnt - a.DiscSum AS Taxexamnt,
        (h.LineAmnt - a.DiscSum) + a.VatSum AS Taxinamnt,

        a.DiscSum,
        a.DiscPrcnt,

        h.LineAmnt - a.DiscSum AS BaseAmount,
        h.LineAmnt - a.DiscSum AS Total1,

        a.VatSum,
        a.DocTotal

    FROM ORIN a
    INNER JOIN OCRD b
        ON a.CardCode = b.CardCode
    LEFT JOIN CRD1 c
        ON b.CardCode = c.CardCode
    INNER JOIN RIN12 d
        ON a.DocEntry = d.DocEntry
    INNER JOIN
    (
        SELECT DISTINCT BaseRef, DocEntry
        FROM RIN1
        WHERE BaseType = '13'
    ) m
        ON a.DocEntry = m.DocEntry
    INNER JOIN
    (
        SELECT COUNT(*) AS [Count], DocEntry, BaseRef
        FROM RIN1
        GROUP BY DocEntry, BaseRef
    ) g
        ON a.DocEntry = g.DocEntry
    INNER JOIN
    (
        SELECT
            SUM(LineTotal) AS LineAmnt,
            DocEntry
        FROM RIN1
        GROUP BY DocEntry
    ) h
        ON a.DocEntry = h.DocEntry
    INNER JOIN
    (
        SELECT
            SUM(VatSum) AS VatSum,
            MAX(Rate) AS Rate,
            DocEntry
        FROM RIN1
        GROUP BY DocEntry
    ) v
        ON a.DocEntry = v.DocEntry
    LEFT JOIN [@I_ZATCA_TAXCODE] t
        ON a.U_ZATCA_TaxCode = t.U_I_Tax_Ex_Code
    CROSS JOIN OADM e
    CROSS JOIN ADM1 f
    WHERE a.DocEntry = @DocEntry
      AND c.AdresType = 'B';
END;
GO


/* =========================================================
   6. @ECREDITMEMO_DETAIL
   ========================================================= */
IF OBJECT_ID('dbo.[@ECREDITMEMO_DETAIL]', 'P') IS NOT NULL
    DROP PROCEDURE dbo.[@ECREDITMEMO_DETAIL];
GO

CREATE PROCEDURE dbo.[@ECREDITMEMO_DETAIL]
    @DocEntry NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        B.ItemCode,
        B.Dscription,
        B.VatSum,
        B.VatGroup,
        ROUND(
            (B.INMPrice /
                CASE
                    WHEN B.Rate = 0.0 THEN 1.0
                    ELSE B.Rate
                END
            ), 2
        ) AS Price,
        B.LineNum,
        CASE
            WHEN A.DocType = 'I' THEN B.Quantity
            WHEN A.DocType = 'S' THEN 1
        END AS Quantity,
        B.U_APIRefNo,
        B.Currency,
        'PCE' AS Unitmsr,
        CASE
            WHEN B.DiscPrcnt > 0 THEN B.PriceBefDi * (B.DiscPrcnt / 100.0)
            ELSE 0
        END AS Discount,
        B.LineTotal AS LineTotal,
        ROUND(B.LineTotal * (C.Rate / 100.0), 2) AS vat,
        B.VatSum AS v,
        ROUND(
            (B.Price /
                CASE
                    WHEN B.Rate = 0.0 THEN 1.0
                    ELSE B.Rate
                END
            ), 2
        ) AS BaseAmount,
        B.DiscPrcnt
    FROM RIN1 B
    INNER JOIN ORIN A
        ON A.DocEntry = B.DocEntry
    INNER JOIN OVTG C
        ON B.VatGroup = C.Code
    WHERE A.DocEntry = @DocEntry;
END;
GO