// Demonstrates the following techniques:
//  Querying metadata using the RetrieveMetadataChanges request
//  See: https://docs.microsoft.com/en-us/dynamics365/customer-engagement/developer/webapi/use-web-api-functions

describe("", function() {
  it("RetrieveMetadataChanges", async function() {
    this.timeout(90000);
    var assert = chai.assert;

    // Query for Account metadata and return only the OptionSet values for the attribute address1_shippingmethodcode
    var request = new class {
      Query = {
        Criteria: {
          Conditions: [
            {
              PropertyName: "LogicalName",
              ConditionOperator: "Equals",
              Value: {
                Value: "account",
                Type: "System.String"
              }
            }
          ],
          FilterOperator: "And"
        },
        Properties: {
          PropertyNames: ["Attributes"]
        },
        AttributeQuery: {
          Properties: {
            PropertyNames: ["OptionSet"]
          },
          Criteria: {
            Conditions: [
              {
                PropertyName: "LogicalName",
                ConditionOperator: "Equals",
                Value: {
                  Value: "address1_shippingmethodcode",
                  Type: "System.String"
                }
              }
            ],
            FilterOperator: "And"
          }
        }
      };

      getMetadata(): any {
        return {
          parameterTypes: {
            AppModuleId: {
              typeName: "Edm.Guid",
              structuralProperty: 1
            },
            ClientVersionStamp: {
              typeName: "Edm.String",
              structuralProperty: 1
            },
            DeletedMetadataFilters: {
              typeName: "mscrm.DeletedMetadataFilters",
              structuralProperty: 3
            },
            Query: {
              typeName: "mscrm.EntityQueryExpression",
              structuralProperty: 5
            }
          },
          operationType: 1,
          operationName: "RetrieveMetadataChanges"
        };
      }
    }();
    var rawResponse = await Xrm.WebApi.online.execute(request);
    var response: {
      DeletedMetadata: Object;
      EntityMetadata: {
        ActivityTypeMask: any;
        Attributes: {
          AttributeOf: any;
          AttributeType: String;
          AttributeTypeName: Object;
          AutoNumberFormat: any;
          CanBeSecuredForCreate: any;
          CanBeSecuredForRead: any;
          CanBeSecuredForUpdate: any;
          CanModifyAdditionalSettings: any;
          ChildPicklistLogicalNames: any[];
          ColumnNumber: any;
          DefaultFormValue: any;
          DeprecatedVersion: any;
          Description: any;
          DisplayName: any;
          EntityLogicalName: any;
          ExternalName: any;
          FormulaDefinition: any;
          HasChanged: any;
          InheritsFrom: any;
          IntroducedVersion: any;
          IsAuditEnabled: any;
          IsCustomAttribute: any;
          IsCustomizable: any;
          IsDataSourceSecret: any;
          IsFilterable: any;
          IsGlobalFilterEnabled: any;
          IsLogical: any;
          IsManaged: any;
          IsPrimaryId: any;
          IsPrimaryName: any;
          IsRenameable: any;
          IsRequiredForForm: any;
          IsRetrievable: any;
          IsSearchable: any;
          IsSecured: any;
          IsSortableEnabled: any;
          IsValidForAdvancedFind: any;
          IsValidForCreate: any;
          IsValidForForm: any;
          IsValidForGrid: any;
          IsValidForRead: any;
          IsValidForUpdate: any;
          LinkedAttributeId: any;
          LogicalName: String;
          MetadataId: String;
          OptionSet: Object;
          ParentPicklistLogicalName: any;
          RequiredLevel: any;
          SchemaName: any;
          SourceType: any;
          SourceTypeMask: any;
        }[];
        AutoCreateAccessTeams: any;
        AutoRouteToOwnerQueue: any;
        CanBeInCustomEntityAssociation: any;
        CanBeInManyToMany: any;
        CanBePrimaryEntityInRelationship: any;
        CanBeRelatedEntityInRelationship: any;
        CanChangeHierarchicalRelationship: any;
        CanChangeTrackingBeEnabled: any;
        CanCreateAttributes: any;
        CanCreateCharts: any;
        CanCreateForms: any;
        CanCreateViews: any;
        CanEnableSyncToExternalSearchIndex: any;
        CanModifyAdditionalSettings: any;
        CanTriggerWorkflow: any;
        ChangeTrackingEnabled: any;
        CollectionSchemaName: any;
        DataProviderId: any;
        DataSourceId: any;
        DaysSinceRecordLastModified: any;
        Description: any;
        DisplayCollectionName: any;
        DisplayName: any;
        EnforceStateTransitions: any;
        EntityColor: any;
        EntityHelpUrl: any;
        EntityHelpUrlEnabled: any;
        EntitySetName: any;
        ExternalCollectionName: any;
        ExternalName: any;
        HasActivities: any;
        HasChanged: any;
        HasFeedback: any;
        HasNotes: any;
        IconLargeName: any;
        IconMediumName: any;
        IconSmallName: any;
        IconVectorName: any;
        IntroducedVersion: any;
        IsAIRUpdated: any;
        IsActivity: any;
        IsActivityParty: any;
        IsAuditEnabled: any;
        IsAvailableOffline: any;
        IsBPFEntity: any;
        IsBusinessProcessEnabled: any;
        IsChildEntity: any;
        IsConnectionsEnabled: any;
        IsCustomEntity: any;
        IsCustomizable: any;
        IsDocumentManagementEnabled: any;
        IsDocumentRecommendationsEnabled: any;
        IsDuplicateDetectionEnabled: any;
        IsEnabledForCharts: any;
        IsEnabledForExternalChannels: any;
        IsEnabledForTrace: any;
        IsImportable: any;
        IsInteractionCentricEnabled: any;
        IsIntersect: any;
        IsKnowledgeManagementEnabled: any;
        IsLogicalEntity: any;
        IsMSTeamsIntegrationEnabled: any;
        IsMailMergeEnabled: any;
        IsManaged: any;
        IsMappable: any;
        IsOfflineInMobileClient: any;
        IsOneNoteIntegrationEnabled: any;
        IsOptimisticConcurrencyEnabled: any;
        IsPrivate: any;
        IsQuickCreateEnabled: any;
        IsReadOnlyInMobileClient: any;
        IsReadingPaneEnabled: any;
        IsRenameable: any;
        IsSLAEnabled: any;
        IsStateModelAware: any;
        IsValidForAdvancedFind: any;
        IsValidForQueue: any;
        IsVisibleInMobile: any;
        IsVisibleInMobileClient: any;
        Keys: any[];
        LogicalCollectionName: any;
        LogicalName: String;
        ManyToManyRelationships: any[];
        ManyToOneRelationships: any[];
        MetadataId: String;
        MobileOfflineFilters: any;
        ObjectTypeCode: any;
        OneToManyRelationships: any[];
        OwnershipType: any;
        PrimaryIdAttribute: any;
        PrimaryImageAttribute: any;
        PrimaryNameAttribute: any;
        Privileges: any[];
        RecurrenceBaseEntityLogicalName: any;
        ReportViewName: any;
        SchemaName: any;
        SyncToExternalSearchIndex: any;
        UsesBusinessDataLabelTable: any;
      }[];
      ServerVersionStamp: String;
    } = await (<any>rawResponse).json();

    assert.equal(
      response.EntityMetadata[0].LogicalName,
      "account",
      "Account metadata returned"
    );
    assert.ok(
      response.EntityMetadata[0].Attributes.length > 0,
      "Account Attributes returned"
    );
  });
});
