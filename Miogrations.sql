-- 1. Необходимо создать таблицу на сервере MSSQL с 3 столбцами

CREATE TABLE data (
    id INT IDENTITY(1,1) PRIMARY KEY,
    date DATE,
    articul NVARCHAR(255),
    sales INT
);


-- 3. Сделать хранимую процедуру на MS SQL сервере для выгрузки данных за определенный период

CREATE PROCEDURE GetReport
    @StartDate DATE,
    @EndDate DATE
AS
BEGIN
    SELECT
        YEAR(date) AS Year,
        MONTH(date) AS Month,
        articul,
        AVG(sales) AS AverageSales,
        SUM(sales) * 100.0 / 
            (SELECT SUM(sales) FROM data WHERE date BETWEEN @StartDate AND @EndDate) 
            AS TotalShare
    FROM
        data WHERE date BETWEEN @StartDate AND @EndDate 
    GROUP BY
        YEAR(date),
        MONTH(date),
        articul
    ORDER BY
        Year,
		MONTH
END