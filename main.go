package main

import (
	"context"
	"fmt"
	"log"
	"os"
	"strconv"
	"strings"
	"unicode"

	"github.com/Azure/azure-sdk-for-go/sdk/azidentity"
	"github.com/Azure/azure-sdk-for-go/sdk/resourcemanager/network/armnetwork/v2"
	"github.com/Azure/azure-sdk-for-go/sdk/resourcemanager/resources/armsubscriptions"
	"github.com/xuri/excelize/v2"
)

var(
	TenantID = os.Getenv("AZURE_TENANT_ID")
	SubscriptionsNameMustContainStr = os.Getenv("SUBSCRIPTION_MUST_CONTAIN_STR")
	SubnetsMap = make(map[string]*Subnets)
)

type AzureAuth struct {
	*azidentity.AzureCLICredential
}
type ResourceID struct{
	subscription string
	rg string
}
type Nsg struct{
	rg string
	nsgName string
}
type Subnets struct{
	subscription string
	rg string
	vnetName string
	subnetName string
	subnetRange string
	ngs map[int]*NSGrules
}
type NSGrules struct{
	priority int32
	name string
	sourceIPAddress []string
	destPorts []string
	action string
}

//getting auth from azure CLI
func NewAzureAuth() *AzureAuth{
	cred, err := azidentity.NewAzureCLICredential(&azidentity.AzureCLICredentialOptions{TenantID: TenantID})
	if err != nil {
		log.Fatal(err)
	}	
	return &AzureAuth{cred}
}

//GET VNET cli
func (azureAuth *AzureAuth) VnetClient(subscriptionID *string) *armnetwork.VirtualNetworksClient{
	clientFactory, err := armnetwork.NewVirtualNetworksClient(*subscriptionID, azureAuth, nil)
	if err != nil {
		log.Fatal(err)
	}
	return clientFactory
}

//GET SUBNET cli
func (azureAuth *AzureAuth) SubnetClient(subscriptionID *string) *armnetwork.SubnetsClient{
	clientFactory, err := armnetwork.NewSubnetsClient(*subscriptionID, azureAuth, nil)
	if err != nil {
		log.Fatal(err)
	}
	return clientFactory
}

func (azureAuth *AzureAuth) NsgClient(subscriptionID *string) *armnetwork.SecurityGroupsClient{
	clientFactory, err := armnetwork.NewSecurityGroupsClient(*subscriptionID,azureAuth,nil)
	if err != nil {
		log.Fatal(err)
	}
	return clientFactory
}


//get Subscription CLI
func (azureAuth *AzureAuth) SubscriptionClient() *armsubscriptions.Client{
	clientFactory, err := armsubscriptions.NewClient(azureAuth, nil)
	if err != nil {
		log.Fatal(err)
	}
	return clientFactory
}

func ResourceIDtoStruct(resourceID *string) ResourceID{
	arr := strings.Split(*resourceID, "/")
	return ResourceID{
		subscription: arr[2],
		rg: arr[4],
	}
}

func GetNsgNameFromResourceID(resourceID *string) *Nsg{
	arr := strings.Split(*resourceID,"/")
	return &Nsg{
		rg: arr[4],
		nsgName: arr[len(arr)-1],
	}
}


func separateLettersAndNumbers(input string) (string, string) {
	numbers:= ""
	letters:=""
	for _, char := range input {
		if unicode.IsLetter(char) {
			letters += string(char)
		} else if unicode.IsDigit(char) {
			numbers += string(char)
		}
	}

	return letters, numbers
}

func NextCell(typeCell string, cell string) string{

	columnToIndexMap := map[string]int{
		"A": 0,
		"B": 1,
		"C": 2,
		"D": 3,
		"E": 4,
		"F": 5,
		"G": 6,
		"H": 7,
		"I": 8,
		"J": 9,
	}
	indexToColumnMap := map[int]string{
		0:"A",
		1:"B",
		2:"C",
		3:"D",
		4:"E",
		5:"F",
		6:"G",
		7:"H",
		8:"I",
		9:"J",
	}

	column, row := separateLettersAndNumbers(cell)
	num, err := strconv.Atoi(row)
	if err != nil {
		fmt.Println("Error during conversion")
	}

	newColumn := column
	newRow := num
	switch typeCell {
		case "column":
			newColumn = indexToColumnMap[columnToIndexMap[column]+1]
		case "row":
			return "A"+strconv.Itoa(num+1)
	}
	
	return newColumn+strconv.Itoa(newRow)
}

func rangeToSubnetName(ipRange string) string{
	
	if SubnetsMap[ipRange] != nil{
		return SubnetsMap[ipRange].subscription+"/"+SubnetsMap[ipRange].vnetName+"/"+SubnetsMap[ipRange].subnetName
	}

	return ipRange
}

func GenerateReport(){

	f := excelize.NewFile()
    defer func() {
        if err := f.Close(); err != nil {
            fmt.Println(err)
        }
    }()

    // Create a new sheet.
    index, err := f.NewSheet("NSGRules")
    if err != nil {
        fmt.Println(err)
        return
    }
	
	cell := "A1"
	for _, subnetItem := range(SubnetsMap){
		
		for _, nsgRule := range(subnetItem.ngs){
			for _, nsgSourceIP := range(nsgRule.sourceIPAddress){

				f.SetCellValue(
					"NSGRules",
					cell, 
					subnetItem.subscription+"/"+subnetItem.vnetName+"/"+subnetItem.subnetName,
				)

				cell = NextCell("column", cell)
				f.SetCellValue(
					"NSGRules",
					cell, 
					nsgRule.name,
				)
				
				cell = NextCell("column", cell)
				f.SetCellValue(
					"NSGRules",
					cell, 
					rangeToSubnetName(nsgSourceIP),
				)

				cell = NextCell("column", cell)
				f.SetCellValue(
					"NSGRules",
					cell, 
					nsgRule.action,
				)
				
				cell = NextCell("column", cell)
				f.SetCellValue(
					"NSGRules",
					cell, 
					strings.Join(nsgRule.destPorts, ","),
				)
				
				cell = NextCell("row", cell)
			}
		}
		
		cell = NextCell("row", cell)
	}


	f.SetActiveSheet(index)
    // Save spreadsheet by the given path.
    if err := f.SaveAs("NSG-QA.xlsx"); err != nil {
        fmt.Println(err)
    }
}




func main(){
	//ctx := context.Background()
	azureAuth := NewAzureAuth()
	
	//listing all subscription from a tenant
	subscriptionList := azureAuth.SubscriptionClient().NewListPager(nil)
	for subscriptionList.More() {
		page, err := subscriptionList.NextPage(context.Background())
		if err != nil {
			log.Fatal(err)
		}

		//list subscriptions
		for _, w := range page.Value {

			if !strings.Contains(*w.DisplayName, SubscriptionsNameMustContainStr){
				continue
			}
			
			//Get all vnets from a subscription
			vnetList := azureAuth.VnetClient(w.SubscriptionID).NewListAllPager(nil)
			for vnetList.More() {

				pageVnet, err := vnetList.NextPage(context.Background())
				if err != nil {
					log.Fatal(err)
				}

				//List Vnets
				for _, valueVnet := range pageVnet.Value {

					resourceID := ResourceIDtoStruct(valueVnet.ID)

					//get all subnet from a vnet
					subnetList := azureAuth.SubnetClient(w.SubscriptionID).NewListPager(resourceID.rg, *valueVnet.Name, nil)
					for subnetList.More() {
						pageSubnet, err := subnetList.NextPage(context.Background())
						if err != nil {
							log.Fatal(err)
						}

						//List subnets
						for _, valueSubnet := range pageSubnet.Value {

							NSGrulesList := make(map[int]*NSGrules)

							//azureAuth.NsgClient(w.SubscriptionID).Get()
							
							if valueSubnet.Properties.NetworkSecurityGroup != nil{

								if valueSubnet.Properties.NetworkSecurityGroup.ID != nil{
									nsgResource := GetNsgNameFromResourceID(valueSubnet.Properties.NetworkSecurityGroup.ID)
									
									nsgProperties, err :=azureAuth.NsgClient(w.SubscriptionID).Get(context.Background(), nsgResource.rg,nsgResource.nsgName, nil)
									if err != nil {
										log.Fatal(err)
									}

									//List of NSG rules
									rulesIndex := 0
									for _,ngsrule := range(nsgProperties.Properties.SecurityRules){

										//ONLY inbounds
										if *ngsrule.Properties.Direction == armnetwork.SecurityRuleDirectionInbound{

											nsgRules := &NSGrules{
												name: *ngsrule.Name,
												priority: *ngsrule.Properties.Priority,
												sourceIPAddress: []string{},
												destPorts: []string{},
												action: "",
											}

											if *ngsrule.Properties.Access == armnetwork.SecurityRuleAccessAllow{
												nsgRules.action = "Allow"
											}else{
												nsgRules.action = "Deny"
											}
											
											//check multiple sources
											if len(ngsrule.Properties.SourceAddressPrefixes) > 0{
												for _,sourceAddresses := range(ngsrule.Properties.SourceAddressPrefixes){
													nsgRules.sourceIPAddress = append(nsgRules.sourceIPAddress, *sourceAddresses)
												}
											}else{
												nsgRules.sourceIPAddress = []string{*ngsrule.Properties.SourceAddressPrefix}
											}
											
											//check multiple ports
											if len(ngsrule.Properties.DestinationPortRanges) > 0{
												for _,destPorts := range(ngsrule.Properties.DestinationPortRanges){
													nsgRules.destPorts = append(nsgRules.destPorts, *destPorts)
												}
											}else{
												nsgRules.destPorts = []string{*ngsrule.Properties.DestinationPortRange}
											}										
											
											NSGrulesList[rulesIndex] = nsgRules
											rulesIndex++
										}
									}
								}
							}

							SubnetsMap[*valueSubnet.Properties.AddressPrefix] = &Subnets{
								subscription: *w.DisplayName,
								rg: resourceID.rg,
								vnetName: *valueVnet.Name,
								subnetName: *valueSubnet.Name,
								subnetRange: *valueSubnet.Properties.AddressPrefix,
								ngs: NSGrulesList,
							}
						}
					}
				}
			}	
		}
	}
	

	GenerateReport()
}