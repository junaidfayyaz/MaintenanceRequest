import * as React from "react";
import { useState, useRef, useEffect } from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../css/style.css";
import {
  IUserProfile,
  IMaintenanceRequestFormData,
} from "../webparts/maintenanceRequest/components/IMaintenanceRequestProps";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Col } from "react-bootstrap";
import AlertModal from "./alertModal/AlertModal";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IRequestUIFormProps {
  context: WebPartContext;
  userprofileAD: IUserProfile;
  EmpId: string;
  onErrorRequiredFields: () => void;
  onSave: (formData: IMaintenanceRequestFormData) => Promise<void>;
  OcpApimKey: string;
  UserRecIdApilink: string;
  Authorization: string;
  CurrentUserEmail?: string;
}

const isAr =
  window.location.pathname.includes("/ar/") ||
  window.location.search.includes("lang=ar");

const MaintenanceRequestUIForm: React.FC<IRequestUIFormProps> = (props) => {
  const [formData, setFormData] = useState<IMaintenanceRequestFormData>({
    requestedBy: props.userprofileAD?.displayName,
    requestedFor: "",
    requestedFor_Title: "",
    requestedFor_key: "",
    Category: "",
    SubCategory: "",
    Department: "",
    Location: "",
    Description: "",
    files: [],
  });

  const [uploadedFiles, setUploadedFiles] = useState<Array<{ name: string; content?: any; progress?: number }>>([]);
  const [showErrorUpload, setShowErrorUpload] = useState("");
  const [showSubmitbtn, setShowSubmitbtn] = useState(false);
  const [errors, setErrors] = useState<{ [field: string]: string }>({});
  const inputRef = useRef<HTMLInputElement>(null);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [showEmptyFileModal, setShowEmptyFileModal] = useState(false);
  const [, setForceUpdater] = useState(0);

  const [categoryOptions, setCategoryOptions] = useState<IDropdownOption[]>([]);
  const [subCategoryOptions, setSubCategoryOptions] = useState<IDropdownOption[]>([]);
  const [listData, setListData] = useState<any[]>([]);

  useEffect(() => {
    _getUserRecId(props.CurrentUserEmail, "requestedBy_key");
    fetchCategories();
  }, []);

  const fetchCategories = async () => {
    try {
      // Access specific site for the list
      const tahyahWeb = Web("https://mcitgovqa.sharepoint.com/sites/Tahyah");
      const items = await tahyahWeb.lists
        .getByTitle("MaintenanceRequestCategoryList")
        .items.getAll();
      
      setListData(items);

      // Extract unique categories
      const uniqueCategories = Array.from(new Set(items.map((item) => item.Category).filter(Boolean)));
      
      setCategoryOptions(
        uniqueCategories.map((cat) => ({ key: cat as string, text: cat as string }))
      );
    } catch (error) {
      console.error("Error fetching categories: ", error);
    }
  };

  const handleCategoryChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      const selectedCategory = option.key as string;
      updateFormDropData(option, "Category");

      const categoryItem = listData.find((item) => item.Category === selectedCategory);
      let categoryKey = "";
      if (categoryItem) {
        // Look for any property name that might contain the Category recId
        const possibleKeys = Object.keys(categoryItem).filter(k => k.toLowerCase().includes("recid") && !k.toLowerCase().includes("sub"));
        if (possibleKeys.length > 0) {
          categoryKey = categoryItem[possibleKeys[0]];
        }
      }
      
      // Fallback to user-provided mapping if not found in SharePoint list
      if (!categoryKey) {
        if (selectedCategory === "Access Control") {
          categoryKey = "87967EC7EA98409C8C2A636924C6C797";
        } else if (selectedCategory === "Health") {
          categoryKey = "34C6EE3A0B354DD599DECDB0624151BE";
        } else if (selectedCategory === "Air Conditioning / Ventilation") {
          categoryKey = "220909EF816F45EAB5C4ED80C48681E3";
        } else if (selectedCategory === "Electrical / Lighting") {
          categoryKey = "EC9986F72483419E99D9F96C9A1B1863";
        } else if (selectedCategory === "Plumbing & Water Systems") {
          categoryKey = "53720C7CE2544D5D90ADD60A3110BC53";
        } else if (selectedCategory === "Interior Fit-out") {
          categoryKey = "E4E283B4BDD24EB192D714D636E2A5B4";
        } else if (selectedCategory === "Fire Systems") {
        categoryKey = "par-D08FBDD8FCD4434B85BEF12F75CF4B59";
      } else if (selectedCategory === "Elevators") {
        categoryKey = "par-DCB8F64B2B274DF1852C51C0DC97147C";
      } else if (selectedCategory === "Cleaning") {
        categoryKey = "par-E1D98ECA66DF4D4F98A625ACE4482589";
      } else if (selectedCategory === "Security") {
        categoryKey = "par-2A2B853272424F679E129BA996561AAC";
      } else if (selectedCategory === "Pest Control") {
        categoryKey = "-8B027D154F64449BBC37690497DBB611";
      } else if (selectedCategory === "Landscaping") {
        categoryKey = "-242A02793B8B4D5BB8FEA98AFCAEA9F5";
      } else if (selectedCategory === "Mail Room") {
        categoryKey = "-6276B7DC33804760B17A9623EDB04806";
      } else if (selectedCategory === "Pantry Service") {
        categoryKey = "-35C4083864304869AD82C9042D97B697";
      } else if (selectedCategory === "Transportation") {
        categoryKey = "-1B9112137E2A4B028551FF05F3DCE00A";
      } else if (selectedCategory === "Meeting Room") {
        categoryKey = "-68701F56FB6C468A8C3C36D98213187B";
      } else if (selectedCategory === "Environment") {
        categoryKey = "-1EEE6900356A4924B3418EAA74B2CE0F";
      } else if (selectedCategory === "WorkSpace") {
        categoryKey = "-C8387403403C467B9CF1F6E3C33C1B8C";
      } else if (selectedCategory === "Quality") {
        categoryKey = "-4FCD70FCF19643D08CB6FDE4CB8D4200";
      } else if (selectedCategory === "Others") {
        categoryKey = "-C06F7C725B924E09ABF2D8EBBB103CA8";
      }
      }

      setFormData((prev) => ({ 
        ...prev, 
        Category_key: categoryKey,
        SubCategory: "",
        SubCategory_key: ""
      }));

      // Filter subcategories based on selected category and ensure uniqueness
      const filteredSubCats = Array.from(
        new Set(
          listData
            .filter((item) => item.Category === selectedCategory)
            .map((item) => item.SubCategory)
            .filter(Boolean)
        )
      );

      setSubCategoryOptions(
        filteredSubCats.map((subcat) => ({ key: subcat as string, text: subcat as string }))
      );
    }
  };

  const handleSubCategoryChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option) {
      const selectedSubCategory = option.key as string;
      updateFormDropData(option, "SubCategory");
      
      const subCategoryItem = listData.find((item) => item.Category === formData.Category && item.SubCategory === selectedSubCategory);
      let subCategoryKey = "";
      if (subCategoryItem) {
        // Look for any property name that might contain the SubCategory recId
        const possibleKeys = Object.keys(subCategoryItem).filter(k => k.toLowerCase().includes("sub") && k.toLowerCase().includes("recid"));
        if (possibleKeys.length > 0) {
          subCategoryKey = subCategoryItem[possibleKeys[0]];
        } else {
           // fallback if it just named recId on the item level
           const fallbackKeys = Object.keys(subCategoryItem).filter(k => k.toLowerCase().includes("recid"));
           if (fallbackKeys.length > 0) {
             // pick one that is not the category key
             const subKey = fallbackKeys.find(k => subCategoryItem[k] !== formData.Category_key);
             if (subKey) subCategoryKey = subCategoryItem[subKey];
           }
        }
      }

      // Fallback to user-provided mapping if not found in SharePoint list
      if (!subCategoryKey) {
        const subCategoryMappings: { [key: string]: string } = {
          "Automated Doors not working": "3E31F1E9EBA1466AB4546AF05F3FC02F",
          "Car Barriers not working": "F8838CF5360E479C87E87C2EE28E4DB4",
          "People Barrier not working": "DC2E58CC674544379E3996C3F3E5AB3B",
          "Roller shutters not working": "8A445EAF3E4847EAAB30269B1D4460B5",
          "Access card not working": "05614D47CE594D499DF6375286C1CEB2",
          "automated door noisy": "6BFF811DE98A4C17B21D38E084DEB1BB",
          "automated door jammed": "0DE05F7CEDB943CDB072AB9D5A36AF9C",
          "roller shutter noisy": "93E91329CC2E492596D7C6B7A59930DB",
          "Roller Shutter PPM Due": "34ABE0BB87464926B3481BC3BD99D554",
          "roller shutter jammed": "E8E0596FF0D04C56B552AB0CBB1599F5",
          "Accidental release of biological agents": "C52A47401AB74224921A9A35240EFAAD",
          "Asbestos Exposure": "183F1CF5193642BB9DBA28A502B2AB64",
          "Ergonomics and Back Injury": "598AF0C0CEBD42B2BECA8D7E82188194",
          "Food Poisoning": "63F99B8195DD42BC9B8FCF8252A47463",
          "Infectious disease": "E512F1AC4A404CCEA14AF7988D4968AC",
          "Malfunction of Self-Contained-Breathing-Apparatus": "E239F094A8B744799DF7048911940BC1",
          "Medical Emergency": "3C9958416A674D179F77C5AF89125E6A",
          "Occupational Skin disease": "387035C6E09542C193932AB4427197AD",
          "Other Health Issue": "14D70B81D3B245F28D5F11853621A81E",
          "Paralysis": "BA6A3598C7EB4F43A33F5599415C2110",
          "Radiation": "5F66098CAA754DE7BE9F90006231E1EA",
          "Swimming Pool Incident": "67F990F905724B70925DD7E81FA6AEC5",
          "Un-hygienic and unhealthy condition": "2C27C46550CC48E5BB4BB04064B4157F",
          "Customer Quality Concern": "952F6817037F449DBADFFA69901A60DD",
          "Electricity failure or breakdown": "0295DEA619E64FA89C6DE8BFD6728530",
          "A/C Too Cold": "67031BCF0FF544F78E648641E173BE03",
          "A/C Not Working": "735874BF09B94231BDF8BA0EE21ACF66",
          "A/C Noisy": "B6245B4B513A4F169D8C227420A62792",
          "A/C Water Leakage": "4079909B57D14E218CB430A4D5699979",
          "A/C Control Not Working": "9550F99B3DFB4244B3AFB0D9D12D5819",
          "A/C Tripping": "1681DC69AB8E498E950A39CD0A68D27B",
          "A/C Bad Odour": "D8D0D90B6B5C45D3AC39558FA847F0F2",
          "A/C Not Cooling": "913D261A1FC44E1AA8D53E56FCF9EC59",
          "A/C Gas Leakage": "211B45F33B04468BB123948BEA3971D8",
          "Thermostat Faulty": "075327A8389D412C8213E3EDBB2AF78B",
          "Ceiling Light Not Working": "8DCEB549436D4A65A35915A79755E1A6",
          "Lamps Fused": "E59E166A551A4FC4A627CF3C01B49DF0",
          "Emergency Light Not Working": "99F675BBBF1C4FA1AE72F4B9962EA68E",
          "Lights Flickering": "B06B0740326B4C55A01FD86F01B03DDA",
          "No Power in the Socket": "6F7CF49A79E349F1B897A95D2632D0BE",
          "Naked Cables": "D0EB4FA8A93845078FA828F78102195C",
          "Exhaust Fan Not Working": "E37A289FE5FD4B78B68653661A85000F",
          "No Power Supply": "514ADACB842B4E928005694E745EA665",
          "Breaker/ELCB/RCD tripping": "71102CFD72FA456C8E28D0ADEE2684B4",
          "Tipani Window Light not working": "3F772D060CC8427CB066789F01D43EEE",
          "Shifting of Light": "B57E7CF1946446169382F3F36CC24427",
          "Additional light required": "BAFF4C1AA1524FF291B7273E5C1D266A",
          "Additional socket required": "453F1F61472246BE88146769F30F7F9F",
          "Light fixture Damaged": "D24D219C3F2A4C3ABC3C3D544DBC34D6",
          "Loose Light Fitting": "CE67EB125CF3424D88C5C42A3BED837B",
          "Irrigation System not working": "C72B1FF50899469BA8033B0C9EE4960E",
          "wash basin/kitchen sink blocked": "ECC1499DAE4B4B84B703CFB49FE1FF5D",
          "no water supply": "9C10B80B12654F23AB6D09058B2FDD54",
          "WC blocked": "8843AF26401C4FBD85B348CB8663F92E",
          "Shataff not working": "EE4ECCB530A34E5487607DE331DD6847",
          "urinal leak": "848B3A520A1845E1990D135F9B157C81",
          "dirty water from taps": "4D5892EDB85E455DB8C0DFF8EA11D4C7",
          "water flooding": "B40906BBFF444D0CA8ABC440EFC07956",
          "low water pressure": "E783489D5A2D4BF49CA34CD6E1A17857",
          "water leakage from ceiling": "C4994D628ADC441BA54AEABE52E64C70",
          "water foul smell": "8627F333FAF74F7FBCD87374E3E8C897",
          "bad odour in the washroom/pantry": "99BF2843D74D4D169BCE03D5F82EFA20",
          "Hand shattaf Damage/not working": "D94F19242B07417FBC326AF0DC5D104D",
          "Water leakage from valve/water tap": "15710A142412497BB5830FA79C29FBD3",
          "water dripping": "D8F5781D09A54CB8A54F0C385B7499FF",
          "broken pipeline": "547CF355F3374A1D8B4E71FD177D0750",
          "Flush not working": "1DA80DA5614B414F9C5FB7024E918961",
          "Faucet loose/damaged": "574BCAED88464FD68AFF673D70695518",
          "Seat cover damaged": "E9CD8681BFA14035B481649D68DD10CE",
          "Tissue holder missing/loose": "8DF3BA8397F04BC59254DB2121456F29",
          "Soap dish not working": "63623E91A8594458AFC00CB70D9B1F31",
          "Interior Floors Repairs ( except tiles)": "1E72119150A744759B0AFD4107279603",
          "Replacement of Floor tiles": "2020A2ED4F2E482EA78DF4CAC22867EC",
          "Interior Ceilings Repairs": "82F4F1B34C504B2EB9772D69E4E24686",
          "Locks Repairs": "444D6CA5BF4C49E5980794CA9206578C",
          "Windows Repairs": "CB902800B6E74FC1931BBA18B3682BE4",
          "paint touch up": "53A7E8C1A4B34EAD8DEE4982641FE4F9",
          "door repairs": "CC98D5D7684740BFB5A413B62617C677",
          "cabinet repairs": "CB66C08CFC844873B8F5E7EDBDAA8523",
          "Indoor Signage Repairs": "A19A54A53FC64906B142F8BD0D827C2C",
          "Outdoor Signage Repairs": "324420A77B6843B1AD44546CE10B964F",
          "Garbage chute door not working": "01ADC2FF183E4B30842D8F2EDAB3E9CB",
          "Furniture Request": "43F1FE0CF13E4BD7B808F73ECEC13812",
          "reporting a fire": "5E9F712AF5464957B6D91FDB2AD7C721",
          "Fire alarm cannot be deactivated": "73576EBB9A984AC893F3924550D2069E",
          "fire extinguisher service": "4558C3D1498D43AEA571A206BE482C63",
          "Spinkler/hose reel leaking": "9C20018FBE964FAFA846497ECF470007",
          "FM 200 System servicing": "1C832641430649B4B0A7F25DBFB608BE",
          "Smoke Detector faulty/dirty": "E790D468918948E4B442C85DF980C372",
          "Fault in Fire Panel": "0EC5A8B895E74D4FA6375F1135BDD6F3",
          "Wet Chemical suppression system needs servicing": "F7E20957DE1F4CCEABDBC7B90CC8C5B5",
          "Third party inspection of Kitchen hood due": "168C91D1730D4F7E9AD12F0744979858",
          "Fire Blanket inHeat detector faulty": "A04443D6E30C41DBAE4DBF353C0B6EE6",
          "spection due": "8981840FFA72432DBF9BDA42EB21460B",
          "Sprinklers loose": "874262B7A9C7413182BAA39403346A46",
          "Paint on Sprinklers Cover": "46F38338920F4E0A8FBD3C68DFE208BC",
          "Hose cabinet damaged": "0F6CEE84F4F045C48B52ACA93C99E14D",
          "Elevators not working": "04544A19DEE740DD819BE6B0F1563A41",
          "Escalators not working": "6215CAC7ACA8494998C73DAA36055580",
          "Elevator jammed": "4CF7EAD0C0934A43B8386566E8EA7189",
          "Elevator button not working": "250800A71610492A8A8384A1A3BE525E",
          "Elevator fan not working": "50578C6DC0254A5F8BF8A0E542A23FAF",
          "Door Noisy": "8578F5FD87AB4118BF727A8B6151E87F",
          "Third party inspection due": "431AD5B03AE34C449BA0D43E17590F9B",
          "False Alarm in 24 by 7 Panel": "0463E20528A04536A0437E5BC283FAE3",
          "Esclator/Elevator PPM Due": "00D981389AC04844BEDFDADFCF22660C",
          "Travelators not working": "CDADB1625C3A4ABABF8A0C83C3E6FB73",
          "Spills": "1420478DD13E4DC2B04656614166D7D5",
          "Cowwebs need to be cleaned": "28285E66D07048F1AF038A9F360DF041",
          "diffusers need to be cleaned": "42972F129B9143A687726E0A9C61AC01",
          "Car parking oil spillage": "0398F3D4144D4F64918BBAF69E1936A1",
          "bad odour": "F5679347E25148D58C72EDDFA78D9E53",
          "window cleaning": "108E3835D3094C1EAB74C713C11D05FC",
          "wash room cleaning": "2561687DDC91401AA87780D2F1F15ABC",
          "pantry needs to be cleaned": "7CAD831266914D199D1FDED7F26F3D7A",
          "dusty surfaces": "D91AC46D3DAC4D809A00FE6D50507DFE",
          "Carpet Shampoo": "66400C38CF024A08813315047717464A",
          "Carpet Dirty": "E91DB1337CC44D789D0328CA5C05FED9",
          "Hangers Dirty": "BB4624E7B5954256814F920782CA753D",
          "Display Shelf Dirty": "752634762491451CB07E1D5044E9ECB7",
          "High level dusting": "78ED7BEABD2440B59C3DA5D11183E52D",
          "Floor Cleaning": "7AA52375B90D4318877E99E5C31961C0",
          "Exhaust fan cleaning": "048BB1E410884CE69EDCE553ECFADBD4",
          "Glass cleaning": "40F38A3EFE0644B6BE83BAB4F277EDA7",
          "Garnden cleaning": "0F0745A970C04A3391C706236B064BDF",
          "External/Surrounding Area cleaning": "C18CC37CD7DD49B88CF0F27A0610F6F4",
          "Toilet cleaning": "D34D0D8828C74D08A59999B36043B2CF",
          "Garbage room not clean": "0D47840391164799A94B9490DB6BA5D4",
          "Man Guards needed": "535F8AF9217347818B78A794EF343AC1",
          "Patrols needed": "14B32898138D48E6A1EE7A209D41FBDA",
          "Access Control not working": "820FF03D3925427D9DC53221CE3548F1",
          "Wrong parking": "1590F5D4F22545859CCF835A038C50BE",
          "Emergency": "6117619F04164B98B0F61160530FF4DF",
          "Pest Control required": "66BED36601C0480AB3D9721AAD6632D8",
          "Cockroach/pest/any/rat/mouse": "51844D682FE0409F9CACA3910B2016FF",
          "weeding needs to be done": "2692D32603994C149DE4188C75DDD242",
          "trimming of trees to be done": "14A214DF58394DAF8022CC9C529C140F",
          "grass needs to be cut": "2F8BFA08F4AA4A3888846300E185D58A",
          "fertilizer needs to be sprayed": "C765F0A7659A4A10848829E7C1FFAE7C",
          "internal plant maintenance": "93EBBFE14D004C2396FAB6075D1DAE0A",
          "external plant maintenance": "68248E0361384D63869CFD99BE236D81",
          "request for incoming mail": "259F9299A144407D8E1191F75B9245BC",
          "request for outgoing mail": "AF2DA585C67846B28346344BB2A475A9",
          "mail delayed": "687FA10CA59E40179A988D4E15472F42",
          "request for pantry service": "9ECFF89CE30E4021ADF75BB81DBE0910",
          "delay in service": "B8884A386B804412B11F30EC6A157B43",
          "Feedback/suggestion": "62557B67A1CB490285F6889C230EC1A3",
          "availability request": "EB9FEEAA45CD45348DC3ABB3BD79883D",
          "safe driving": "D338DDF5BD8A4E059647602D34A355DF",
          "vehicle cleanliness": "933C226FE3EF46C3B8125A60E4E25E8F",
          "driver behaviour": "C52C37466743452C99558C6B3150A106",
          "late arrival": "32330EDDEF6F4BF4BED915B66D18A062",
          "route request": "C220ADB6C5964B0CB0F97C19B9CB9FDB",
          "Meeting room set up": "981AB5A7C6EC414B9FA51DF71F9A12FF",
          "projector request": "A53D04F0D606469B9DE5C6D552CFF45B",
          "projector malfunction": "4E5354D2746A46C1998966E9C95CB81D",
          "telephone malfunction": "DBBE293EE32F4A6DAAC1742E5FB08DA2",
          "cleaning concernsv": "2056E898E3494167AD1AAA8B5B9E4948",
          "miscellaneous (specify)": "40CC8D9F8EEB49209F8E1E345CB5275D",
          "Contaminated water": "18E1D64BBCBA4DDB9673F29362651944",
          "Environmental Hazard": "EE6BCEDF31A4431CA0B2E21DC3FFA476",
          "Environmental Incident": "8CCB99BE3F9E4987A1E7767DBF887028",
          "Flammable liquid spill": "029B26FC3B274F39BADFC72D66ADD5BF",
          "High Noise": "1285E84DAD104CFDA367C4AD6BBBC66C",
          "Poor Lighting": "363F67CA48E340478C1D1D9C867096EC",
          "Sewage Water release": "01C38F9D99A54C53A5C22CDC269BCEBE",
          "Toxic Chemical spill": "778DC5DDFA164771AB17779F956A63AB",
          "Toxic gas release": "2C5886580F6C4D92A4F7B834F69F36F0",
          "Vibration": "B74C376A43494FF08852FC242138FC1B",
          "New Joiner": "C7EC70928F4D4F3686A6F433819A1B5E",
          "Relocating": "71D87AD2543C4E96BA0959DFCBFEA83E",
          "Partitions": "29B9ECBF9DE9437AB2C580D726CD3AD0",
          "Others": "7065ECB1F62E4781B40104AA7F814569",
          "Equipment failure or breakdown": "140272D94B0246979033ED50D3E43501",
          "Indoor Air Quality": "F0773D8A16354503BB619F65458D5822",
          "Property Damage": "DBBAB81295AD44C7ADE07E971F66E6B8",
          "Quality Incident": "14248E3044934E09B24D919BDFD0B719",
          "Re-work and repeated problem": "64825D46B218432496C2356F05B06F33",
          "Other Quality Concern": "8D79136D93E04A5184D9EF60FF80898A",
          "Defective Hand Tool": "3CD7A68BFBE84CC29594C6F422762F70",
          "Defective Ladder or scaffold": "3236A079DE20498A886538F6BE050DD6",
          "Electric Shock": "A33CE224FDF2414E807B6AFA052BBF77",
          "electrical circuit overloaded": "FFEF85E7CE494EC6A7020CA6EEC4DC86",
          "Electrical short circuit": "AA23A2F280C2400A8C0A28DC51F993E4",
          "Elevator or Escalator failure": "1CE04FDA30B24C818AF0C6DA1FD56D05",
          "Excavation Incident": "F23C7A75147F498FAC0E377E2C60E4BA",
          "Expired Fire Extinguisher": "B0E478469DF74A348925C781E58FE77A",
          "Explosion or pressure vessel incident": "DEE77E237EA84F5C92D943F3554F8620",
          "Exposed Electrical wire or panel": "2F17A35DA4D3405286BAC87D60192D9D",
          "FAC - First Aid Case": "969C60E29EA34E029DFC38F7BB66BBAF",
          "False Fire Alarm": "D8EFE76959A74F67B4843E70E3F0D33F",
          "Fire Incident": "519649D09FEB48838B20BAB6BE1F416F",
          "Fire Panel out of order or showing fault": "68B29A784AB14BB2BBA126137173E663",
          "Fracture": "8C0D96163630452DA8BE95589EE955F4",
          "Heat Stress": "8A3AA21C47A14E88A245EBB380CEBDE5",
          "MTC - Medical Treatment Case": "F0DABA1689DF4268AFF15ABD4EC5C031",
          "NMI - Near Miss Incident": "1EA130FC3D8F4B9299395FF6B571001E",
          "Power Failure": "E6B5CA82FAC0453DBDA1762DD8165DE2",
          "Reptile bite": "B921A6D0EEA342B99DE635E3B9F1CD1B",
          "Sabotage": "15E8B554C3634D97B5E8C68ED9B0A8AB",
          "Scaffolding Collapse": "07B26EDB305E4E29992EE1BB1DF0E428",
          "Sexual Harassmen": "FBB62F11D02D430583730A64FC598950",
          "Strike, people not at work": "049EEBC6B9FF4E42BCE00003DD0BAAF3",
          "Suffocation": "EEB800C30A34452C975830896E529752",
          "Suicide or self-harm": "08633E60908C42F8A9CED49FAC72F7A6",
          "Suspended equipment incident": "8BE2C7484C9D4DE6BBF74003B40AAE40",
          "Suspicious bag, Packet and letter": "162470A93E924591AD5F75B0ABC8FD24",
          "Unsafe Act": "D9004F37181E4B2990C5BF9E59E5459E",
          "Unsafe Condition": "57F98ECA763D4A26BA7463A18F5C254F",
          "Vehicle Incide": "3089EA37B0E84408BDDE1D3B856B38FF",
          "Violence at work or fighting": "DC9BDDE83F424F9A9191E7226D924D7A",
          "Water supply interruption": "8F38F91E3D224CD78C907B1AA2B982A4"
        };
        if (selectedSubCategory === "Furniture Request" && formData.Category === "WorkSpace") {
          subCategoryKey = "BAC5DF2A75A94192A0B3EC7E07355920";
        } else if (selectedSubCategory === "Others" && formData.Category === "Others") {
          subCategoryKey = "C43E74B524E747C7960DE63632E3C499";
        } else if (subCategoryMappings[selectedSubCategory]) {
          subCategoryKey = subCategoryMappings[selectedSubCategory];
        }
      }

      setFormData((prev) => ({ ...prev, SubCategory_key: subCategoryKey }));
    }
  };

  function handleInputChange(field: string, value: any) {
    setFormData((prev) => ({ ...prev, [field]: value }));
  }

  function validateForm() {
    const newErrors: { [field: string]: string } = {};

    if (!formData.requestedFor.trim())
      newErrors.requestedFor = isAr ? "متلقي الخدمة مطلوب" : "RequestedFor is required";

    if (!formData.Category.trim())
      newErrors.Category = isAr ? "الفئة مطلوبة" : "Category is required";

    if (!formData.SubCategory.trim())
      newErrors.SubCategory = isAr ? "الفئة الفرعية مطلوبة" : "SubCategory is required";

    if (!formData.Department.trim())
      newErrors.Department = isAr ? "القسم مطلوب" : "Department is required";

    if (!formData.Description.trim())
      newErrors.Description = isAr ? "الوصف مطلوب" : "Description is required";

    setErrors(newErrors);

    if (Object.keys(newErrors).length > 0) {
      props.onErrorRequiredFields();
      return false;
    }
    return true;
  }

  async function handleSubmit() {
    setErrors({});
    if (!validateForm()) return;
    setShowSubmitbtn(true);
    await props.onSave(formData);
    setShowSubmitbtn(false);
  }

  const displayName = props.userprofileAD?.displayName;
  let initials = "";
  if (displayName && displayName.trim()) {
    const parts = displayName.split(" ");
    initials = parts[0][0] + parts[parts.length - 1][0];
  }

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory as any,
    spHttpClient: props.context.spHttpClient as any,
  };

  const _getPeoplePickerItems = async (
    selectedUserProfiles: any[],
    internalName: string,
    internalName_text: string,
    internalName_key: string
  ) => {
    if (selectedUserProfiles.length > 0) {
      const emails = selectedUserProfiles[0].id.split("|")[2];
      const title = selectedUserProfiles[0].text;
      const updatedErrors = { ...errors };
      if (title.trim()) {
        delete updatedErrors[internalName];
      }
      setErrors(updatedErrors);
      handleInputChange(internalName, emails);
      handleInputChange(internalName_text, title);
      _getUserRecId(emails, internalName_key);
    } else {
      handleInputChange(internalName, "");
      handleInputChange(internalName_text, "");
      handleInputChange(internalName_key, "");
    }
  };

  const _getUserRecId = async (email: string | undefined, columnkey: string) => {
    if (!email) return;
    try {
      const response = await fetch(props.UserRecIdApilink, {
        method: "GET",
        headers: {
          "Ocp-Apim-Subscription-Key": props.OcpApimKey,
          Authorization: props.Authorization,
          Email: email,
        },
      });
      if (response.ok) {
        const rawResponse = await response.text();
        const jsonStart = rawResponse.indexOf("{");
        if (jsonStart !== -1) {
          const jsonString = rawResponse.slice(jsonStart);
          const parsedData = JSON.parse(jsonString);
          let RecId = parsedData.value[0].RecId;
          handleInputChange(columnkey, RecId);
        }
      }
    } catch (error: any) {
      console.error("Error getting UserRecId:", error);
    }
  };

  const updateFormData = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue: string | undefined,
    column: any
  ) => {
    const value = newValue ?? "";
    setFormData((prev) => ({
      ...prev,
      [column]: value,
    }));

    setErrors((prevErrors) => {
      const newErrors = { ...prevErrors };
      if (newErrors[column] && value.trim() !== "") {
        delete newErrors[column];
      }
      return newErrors;
    });

    forceUpdate();
  };

  const updateFormDropData = (option: any, column: any) => {
    setFormData((prev) => ({ ...prev, [column]: option?.key as string }));
    setErrors((prevErrors) => {
      const newErrors = { ...prevErrors };
      if (newErrors[column] && option.key) {
        delete newErrors[column];
      }
      return newErrors;
    });
    forceUpdate();
  };

  const forceUpdate = () => setForceUpdater((prev) => prev + 1);

  const removeAttachment = (fileName: string) => {
    const updatedFile = uploadedFiles.filter((file) => file.name !== fileName);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setUploadedFiles(updatedFile);
    handleInputChange("files", updatedFile);
  };

  const getRandomNumber = (min: number = 1, max: number = 1000): number => {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  };

  function updateFileName(file: File, newName: string) {
    return new File([file], newName, {
      type: file.type,
      lastModified: file.lastModified,
    });
  }

  const randomNumber = getRandomNumber(1, 1000);

  const readFile = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    const requesterFileList = e.target.files;
    if (requesterFileList && requesterFileList.length > 0) {
      const file = requesterFileList[0];
      const reader = new FileReader();
      reader.onload = () => {
        const fileContent = reader.result;
        if (fileContent === "") {
          setShowEmptyFileModal(true);
          e.target.value = "";
        } else {
          setShowEmptyFileModal(false);
          const fileExtension = file.name.substring(file.name.lastIndexOf(".") + 1, file.name.length);
          const fileName =
            file.name
              .substring(0, file.name.lastIndexOf(".") + 1)
              .replace(/[&\/\\#~%":*. [\]!¤+`´^?<>|{}]/g, "") +
            randomNumber +
            "." +
            fileExtension;
          
          const updatedFile = updateFileName(file as File, fileName);
          const newFile = {
            name: fileName,
            content: updatedFile,
            progress: 0
          };
          
          setUploadedFiles((prevFiles) => {
            const updatedFiles = [...prevFiles, newFile];
            setFormData((prev) => ({ ...prev, [field]: updatedFiles }));
            return updatedFiles;
          });
          
          let currentProgress = 0;
          const interval = setInterval(() => {
            if (currentProgress >= 100) {
              clearInterval(interval);
            } else {
              currentProgress += 10;
              setUploadedFiles((prevFiles) =>
                prevFiles.map((f) =>
                  f.name === newFile.name ? { ...f, progress: currentProgress } : f
                )
              );
            }
          }, 300);
        }
      };

      reader.onerror = (error) => {
        alert("Error reading the file: " + error);
      };
      reader.readAsText(file as Blob);
    }
  };

  return (
    <div>
      {showEmptyFileModal && (
        <AlertModal
          showModal={true}
          handleShowModal={() => setShowEmptyFileModal(true)}
          handleCloseModal={() => setShowEmptyFileModal(false)}
          heading={"Warning"}
          message={"This attachment has no content"}
          style={""}
          section={"rejected"}
          icon={"WarningSolid"}
          isAr={isAr}
        />
      )}
      <div className="maincontainer">
        <div className="header-top">
          <div className="person-image">{initials}</div>
          <div>
            <div className="person-name">{props.userprofileAD?.displayName}</div>
            <div className="person-description">
              {props.userprofileAD?.jobTitle} | ID: {props.EmpId ? props.EmpId : "N/A"}
            </div>
          </div>
        </div>
        <div className="textContainer">
          <h2 className="form-heading">
            {isAr
              ? "يرجى تعبئة جميع الحقول المطلوبة لتسهيل اعتماد الطلب"
              : "Kindly complete the form below"}
          </h2>

          <div className="fieldContainer">
            {/* Requested By */}
            <TextField
              type="text"
              label={isAr ? "تم التقديم بواسطة" : "Requested By"}
              className="form-text"
              readOnly
              value={props.userprofileAD?.displayName || ""}
            />
            
            {/* Requested For */}
            <div className={`people-picker-wrapper ${errors.requestedFor ? "error-border" : ""}`}>
              <PeoplePicker
                context={peoplePickerContext}
                titleText={isAr ? "مطلوب لـ *" : "Requested For *"}
                personSelectionLimit={1}
                groupName={""}
                groupId=""
                defaultSelectedUsers={[formData.requestedFor]}
                showtooltip={true}
                disabled={false}
                searchTextLimit={1}
                onChange={(e) => {
                  _getPeoplePickerItems(e, "requestedFor", "requestedFor_Title", "requestedFor_key");
                }}
                principalTypes={[PrincipalType.User]}
                resolveDelay={200}
              />
            </div>

            {/* Category Dropdown */}
            <Dropdown
              label={isAr ? "الفئة *" : "Category *"}
              placeholder={isAr ? "حدد الفئة" : "Select Category"}
              selectedKey={formData.Category}
              className={`dropdownfield ${errors.Category ? "error-field" : ""}`}
              styles={{
                root: {
                  marginBottom: '0px'
                },
                dropdown: {
                  borderColor: errors.Category ? "red" : undefined,
                  height: '60px'
                },
                title: {
                  height: '60px',
                  display: 'flex',
                  alignItems: 'center',
                  padding: '0 12px',
                  borderRadius: '6px',
                  borderColor: '#ccc',
                  color: !formData.Category ? '#999' : undefined
                },
                caretDownWrapper: {
                  height: '60px',
                  lineHeight: '60px'
                }
              }}
              onChange={handleCategoryChange}
              options={categoryOptions}
            />

            {/* SubCategory Dropdown */}
            <Dropdown
              label={isAr ? "الفئة الفرعية *" : "SubCategory *"}
              placeholder={isAr ? "حدد الفئة الفرعية" : "Select SubCategory"}
              selectedKey={formData.SubCategory}
              disabled={!formData.Category}
              className={`dropdownfield ${errors.SubCategory ? "error-field" : ""}`}
              styles={{
                root: {
                  marginBottom: '0px'
                },
                dropdown: {
                  borderColor: errors.SubCategory ? "red" : undefined,
                  height: '60px'
                },
                title: {
                  height: '60px',
                  display: 'flex',
                  alignItems: 'center',
                  padding: '0 12px',
                  borderRadius: '6px',
                  borderColor: '#ccc',
                  color: !formData.SubCategory ? '#999' : undefined
                },
                caretDownWrapper: {
                  height: '60px',
                  lineHeight: '60px'
                }
              }}
              onChange={handleSubCategoryChange}
              options={subCategoryOptions}
            />

            {/* Department Textbox */}
            <TextField
              label={isAr ? "القسم *" : "Department *"}
              placeholder={isAr ? "أدخل القسم" : "Enter Department"}
              value={formData.Department}
              onChange={(ev, newValue) => updateFormData(ev, newValue, "Department")}
              className={`form-text ${errors.Department ? "error-field" : ""}`}
              styles={{
                field: { '::placeholder': { color: '#999', opacity: 1 } }
              }}
            />

            {/* Location Textbox (Optional) */}
            <TextField
              label={isAr ? "الموقع" : "Location"}
              placeholder={isAr ? "أدخل الموقع" : "Enter Location"}
              value={formData.Location}
              onChange={(ev, newValue) => updateFormData(ev, newValue, "Location")}
              className="form-text"
              styles={{
                field: { '::placeholder': { color: '#999', opacity: 1 } }
              }}
            />
          </div>

          <div className="description_div">
            {/* Description Multi-line */}
            <TextField
              label={isAr ? "الوصف *" : "Description *"}
              placeholder={isAr ? "أدخل الوصف" : "Enter Description"}
              value={formData.Description}
              multiline
              rows={4}
              type="text-area"
              className={`text-area ${errors.Description ? "error-field" : ""}`}
              onChange={(ev, newValue) => updateFormData(ev, newValue, "Description")}
              styles={{
                root: { color: "#555" },
                fieldGroup: { border: "1px solid #ccc", borderColor: errors.Description ? "red" : undefined },
                field: { color: "#555", '::placeholder': { color: '#999', opacity: 1 } },
              }}
            />
          </div>

          <Col className="mt-4">
            <div style={{ display: "flex", alignItems: "end" }}>
              <label
                className="attach_label"
                style={{
                  marginRight: "4px",
                  marginTop: "24px",
                  fontSize: "14px",
                  fontFamily: "Segoe UI",
                  color: "#555",
                  fontWeight: "600",
                  marginBottom: "11.5px",
                }}
              >
                {isAr
                  ? "يرجى إرفاق مستندات أو صور توضيحية (اختياري)"
                  : "Please attach any supporting documents or pictures (optional)."}
              </label>
            </div>

            <div className="attachment-container">
              <div className="attachment-placeholder">
                <img
                  className="attachment-icon"
                  src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACgAAAAoCAYAAACM/rhtAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAKoSURBVHgB7ZnbjdpAFIaPDYLX7WDdQUBcJJ4CFWRTQUgHbAVABbAVQCpIOljyhIS4leAO4rwB4pL/sMcbxxdij8dWFO0vjcY7c8Z8mjOeOXPWoAgtFotKsVgc4bGCckfpZO/3+06r1bIpoYywxuVy+WCa5kTAHJSfpK57qZUgA4Cbzca6XC7PeOT6CaCDarXqkKLW67WdBtL0N5zP567ATWu1Wi8NnE9bfm+5XH7GJMReMgFAwzDeXztM8wtpFN7bIYFkD8WFNKM6drudTRrFnhBIG6USF9KkHBUC+fVvY3IFZAHS9kC2V6vV5JZ9boDz+dxyn72QqLu3IDMHhBu/cV0qlfre9hDIUdj4ImUs7AZjQH5gCOyJbXpx7VVof7VDfw+QDra24R/jKWN5Zuo7ioXS9hWLfkN2/eMzn0EWQzKMbCuVMBs5vQLSCggX9VDdw02PYf1yKs3C+uD+sGZ9LgZcHy4ayVqakCZpARS4gfu3fJV90qDUgD44N7DgE2OgAzIVIOLGTy4cFvlnkrjxeDx2dEEqA0pQOxW4R3wYU7ev2WxuT6fTR34WyB4pSgmQrwMScTPcEHBjv02j0ZjJrDLkiGebFKQEWCgUHlDdCdwgyo5nlWeXXiCrpCAlQIbiH7wF57Ed80nCVwdSkPJGjU13m8B2RorKPR5MqjfAtPrnAXWHWz9QLqRRWgElMNUqrYAasxCv+r/WoBz67yiFcPRt6/X6U1z7RIBYY3zgVyiF8A4bVWaAHEJZlE6xj0hWIkC5ndmUo8I+kmtUjPSvRTkJ11F32QR2gQAgFvGMa8R8Wi49cYTf7EkdcH8AUMJ4myTzxNEzZSROKElOhj8+53A4DP02oUl0b56a8pHDFy2+y/g7jKgRAjkgPf+GiBRnvzBz46jE+i8JiDR7F2tlUAAAAABJRU5ErkJggg=="
                  alt="Attachment Icon"
                />
                {isAr
                  ? "يرجى إرفاق مستندات أو صور توضيحية (اختياري)"
                  : "Please attach any supporting documents or pictures (optional)"}

                <input
                  type="file"
                  ref={fileInputRef}
                  multiple={true}
                  onChange={(e) => {
                    readFile(e, "files");
                  }}
                />
              </div>
              <span style={{ color: "red" }}>
                {errors.files || showErrorUpload}
              </span>
            </div>

            {uploadedFiles.map((file, index) => (
              <div key={index}>
                <div className="uploadeditems">
                  <strong>{file.name}</strong>
                  <div className="progresscontainer">
                    <div
                      className="progressbar"
                      id="progressbar"
                      style={{ width: `${file.progress || 0}%` }}
                    ></div>
                  </div>
                  <div
                    className="cancelbtn"
                    onClick={() => {
                      removeAttachment(file.name);
                    }}
                  >
                    X
                  </div>
                </div>
              </div>
            ))}
          </Col>

          <div className="buttonContainer">
            <PrimaryButton
              disabled={showSubmitbtn}
              onClick={() => {
                handleSubmit();
              }}
              styles={{ root: { fontSize: "20px" } }}
              text={isAr ? "تقديم" : "Submit"}
              className="submit-formbtn"
            />
            <DefaultButton
              text={isAr ? "مسح" : "Clear"}
              className="cancel-formbtn"
              onClick={() => {
                setFormData({
                  requestedBy: props.userprofileAD?.displayName,
                  requestedFor: "",
                  requestedFor_Title: "",
                  requestedFor_key: "",
                  Category: "",
                  Category_key: "",
                  SubCategory: "",
                  SubCategory_key: "",
                  Department: "",
                  Location: "",
                  Description: "",
                  files: [],
                });
                setUploadedFiles([]);
                setShowErrorUpload("");
                setErrors({});
                if (inputRef.current) inputRef.current.value = "";
              }}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default MaintenanceRequestUIForm;
