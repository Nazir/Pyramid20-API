# README #

| Project name              | Project date |
| ------------------------- | ------------ |
| MeterPoint.withoutNPandLP | 21.12.2020   |

Information
-----------
Реализовано при помощи API «Пирамида 2.0» http://psdn.sicon.ru/

### Выводимые данные:
1. Филиал
2. РЭС
3. Наименование точки учёта
4. Тип счётчика
5. Номер счётчика

### Условия по точкам учёта:
1. Содержащие атрибут «Источник затрат установки ПУ» = «Энергосервисный договор 2019»
2. Не привязанные к справочникам ФЛ или ЮЛ
3. Расход за последние 6 месяцев не больше 1 кВт*ч

Installation
------------

`customXml\item1.xml`:
```xml
<PyramidReport>
    <Parameters>
        <Parameter>
            <Id>ClassifierItems</Id>
            <ParameterName>Элемент классификатора</ParameterName>
            <DataType>69a2cc6a-ae31-43be-b44e-ad9fbc966802</DataType>
            <IsArray>True</IsArray>
            <Value>
                <Item>04f95539-75e4-4ef0-9670-2a00ccb7e316</Item>
                <Item>8346f4ea-8db8-4793-b8a6-aec049dd96fd</Item>
                <Item>a0723917-6fef-4bd1-840d-0444ba5484ed</Item>
                <Item>af002148-8d84-40c5-93f1-d7975b01826b</Item>
                <Item>3d1c37d5-d269-4c81-a7ff-113201667d7d</Item>
                <Item>72e12eb5-26b6-4cff-8ee9-9a54970b91fc</Item>
                <Item>2bdf269f-da03-4e21-be8a-48d36245f6d2</Item>
                <Item>6d2e5001-7a9f-4863-b857-49b89ad1f780</Item>
                <Item>ce8d7e89-f84b-41b8-92c2-811594c5a633</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier>5bfa0ed5-b1b2-46a8-b50d-41c02dbf4884</Classifier>
        </Parameter>
        <Parameter>
            <Id>isMeterMt880</Id>
            <ParameterName>Это MT880?</ParameterName>
            <DataType>Boolean</DataType>
            <IsArray>False</IsArray>
            <Value>
                <Item>False</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
        <Parameter>
            <Id>isMeterAm550</Id>
            <ParameterName>Это AM550?</ParameterName>
            <DataType>Boolean</DataType>
            <IsArray>False</IsArray>
            <Value>
                <Item>True</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
        <Parameter>
            <Id>startDate</Id>
            <ParameterName>Начало периода</ParameterName>
            <DataType>DateTime</DataType>
            <IsArray>False</IsArray>
            <Value>
                <Item>2020-06-01T00:00:00.0000000+03:00</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
        <Parameter>
            <Id>stopDate</Id>
            <ParameterName>Конец периода</ParameterName>
            <DataType>DateTime</DataType>
            <IsArray>False</IsArray>
            <Value>
                <Item>2020-12-02T00:00:00.0000000+03:00</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
        <Parameter>
            <Id>monthsCount</Id>
            <ParameterName>Количество месяцев</ParameterName>
            <DataType>Integer</DataType>
            <IsArray>False</IsArray>
            <Value>
                <Item>6</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
        <Parameter>
            <Id>attr1</Id>
            <ParameterName>Источник затрат установки ПУ</ParameterName>
            <DataType>String</DataType>
            <IsArray>False</IsArray><Value>
            <Item>Энергосервисный договор 2019</Item>
            </Value><IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
        <Parameter>
            <Id>maxDataCount</Id>
            <ParameterName>Максимальное количество данных</ParameterName>
            <DataType>Integer</DataType>
            <IsArray>False</IsArray>
            <Value>
                <Item>0</Item>
            </Value>
            <IsVisible>True</IsVisible>
            <Classifier></Classifier>
        </Parameter>
    </Parameters>
    <Script>
    </Script>
    <Helpers/>
</PyramidReport>
```
