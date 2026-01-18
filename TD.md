graph TD
    Start([User Opens Spreadsheet]) --> Menu[onOpen: Create Menu]
    Menu --> Choice{User Selects Action}
    
    Choice -->|Process Next 50| Batch[processBatch]
    Choice -->|Process All Empty| All[processAllEmpty]
    Choice -->|Show Statistics| Stats[showStatistics]
    Choice -->|Reprocess Errors| Errors[reprocessErrors]
    
    Batch --> ReadData[Read All Sheet Data<br/>Single Operation]
    All --> Confirm{User Confirms?}
    Confirm -->|No| End([End])
    Confirm -->|Yes| ReadData
    
    ReadData --> FindEmpty[findEmptyRows<br/>Identify rows with empty results]
    FindEmpty --> CheckRows{Rows Found?}
    CheckRows -->|No| NoRows[Show 'No rows to process']
    CheckRows -->|Yes| ProcessLoop[processRows Loop]
    
    ProcessLoop --> Cache{Category Cache<br/>Exists?}
    Cache -->|No| GetCategories[getCountryCategories<br/>Read country sheet]
    Cache -->|Yes| UseCache[Use Cached Data]
    GetCategories --> CacheStore[Store in categoryCache]
    UseCache --> Validate
    CacheStore --> Validate
    
    Validate{Valid Data?} -->|No| ErrorMsg[Add Error to Results]
    Validate -->|Yes| Classify[classifyProduct]
    
    Classify --> BuildPrompt[buildPrompt<br/>Create AI prompt]
    BuildPrompt --> CallAPI[callGeminiAPI]
    
    CallAPI --> Retry{API Success?}
    Retry -->|No, Retry < 3| Wait[Sleep & Retry<br/>Exponential Backoff]
    Wait --> CallAPI
    Retry -->|No, Max Retries| APIError[Throw Error]
    Retry -->|Yes| ParseResponse[Parse JSON Response]
    
    ParseResponse --> TieBreaker[applyTieBreakerLogic]
    
    TieBreaker --> HighConf{Confidence ≥ 0.85?}
    HighConf -->|Yes| ReturnCat1[Return Category 1]
    HighConf -->|No| LowConf{Confidence < 0.5?}
    
    LowConf -->|Yes| ManualReview[Return 'Manual Review']
    LowConf -->|No| ComparePerc{Compare<br/>Percentages}
    
    ComparePerc -->|Cat2 % > Cat1 %| ReturnCat2[Return Category 2<br/>with Low Confidence Flag]
    ComparePerc -->|Cat1 % ≥ Cat2 %| ReturnCat1
    
    ReturnCat1 --> AddResult[Add Result to Array]
    ReturnCat2 --> AddResult
    ManualReview --> AddResult
    ErrorMsg --> AddResult
    APIError --> AddResult
    
    AddResult --> MoreRows{More Rows<br/>to Process?}
    MoreRows -->|Yes| RateLimit[Sleep 100ms<br/>Rate Limiting]
    RateLimit --> ProcessLoop
    MoreRows -->|No| WriteResults[writeResults<br/>Batch Write to Sheet]
    
    WriteResults --> Summary[showProcessingSummary<br/>Display Stats]
    Summary --> End
    NoRows --> End
    
    Errors --> ReadData
    Stats --> CountStats[Count classified,<br/>unclassified, errors]
    CountStats --> ShowStats[Display Statistics Dialog]
    ShowStats --> End
    
    style Start fill:#e1f5e1
    style End fill:#ffe1e1
    style CallAPI fill:#fff4e1
    style WriteResults fill:#e1f0ff
    style TieBreaker fill:#f0e1ff
    style ProcessLoop fill:#ffe1f0
