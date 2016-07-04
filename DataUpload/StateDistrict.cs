using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload
{
    public class StateDistrict
    {
        public string State { get; private set; }
        public List<string> Districts { get; private set; }
        public  List<StateDistrict> AllStatesWithDistricts
        {
            get
            {
                return new List<StateDistrict> {


new StateDistrict {
                        State = "N/A",
                        Districts = new List<string> {"N/A"}
                        },
new StateDistrict {
                        State = "TELANGANA",
                        Districts = new List<string> {"Adilabad","Warangal","Karim Nagar","Mahabub Nagar","K.V.Rangareddy","Medak","Nalgonda","Nizamabad","Hyderabad","Khammam"}
                        },
new StateDistrict {
                        State = "ANDHRA PRADESH",
                        Districts = new List<string> {"Ananthapur","Cuddapah","Chittoor","Kurnool","Prakasam","West Godavari","Krishna","Nellore","Guntur","East Godavari","Visakhapatnam","Vizianagaram","Srikakulam"}
                        },
new StateDistrict {
                        State = "PONDICHERRY",
                        Districts = new List<string> {"Pondicherry","Mahe","Karaikal"}
                        },
new StateDistrict {
                        State = "ASSAM",
                        Districts = new List<string> {"Lakhimpur","Dibrugarh","Dhemaji","Marigaon","Nagaon","Karbi Anglong","Golaghat","Sibsagar","Jorhat","Tinsukia","Karimganj","Hailakandi","Cachar","North Cachar Hills","Darrang","Sonitpur","Bongaigaon","Kokrajhar","Goalpara","Dhubri","Kamrup","Nalbari","Barpeta"}
                        },
new StateDistrict {
                        State = "BIHAR",
                        Districts = new List<string> {"Begusarai","Khagaria","Darbhanga","Madhubani","Muzaffarpur","Samastipur","Sitamarhi","East Champaran","Supaul","Vaishali","Purnia","Araria","Katihar","Kishanganj","Madhepura","Saharsa","Saran","Siwan","Bhojpur","Sheohar","Gopalganj","West Champaran","Gaya","Aurangabad(BH)","Arwal","Bhagalpur","Banka","Munger","Buxar","Jehanabad","Patna","Sheikhpura","Jamui","Lakhisarai","Nalanda","Nawada","Rohtas","Kaimur (Bhabua)"}
                        },
new StateDistrict {
                        State = "CHATTISGARH",
                        Districts = new List<string> {"Kanker","Bastar","Dantewada","Bijapur(CGH)","Narayanpur","Bilaspur(CGH)","Janjgir-champa","Korba","Durg","Rajnandgaon","Kawardha","Surguja","Raigarh","Jashpur","Koriya","Raipur","Mahasamund","Dhamtari","Gariaband"}
                        },
new StateDistrict {
                        State = "DELHI",
                        Districts = new List<string> {"East Delhi","North East Delhi","North West Delhi","North Delhi","Central Delhi","New Delhi","South Delhi","South West Delhi","West Delhi"}
                        },
new StateDistrict {
                        State = "GUJARAT",
                        Districts = new List<string> {"Ahmedabad","Gandhi Nagar","Banaskantha","Mahesana","Surendra Nagar","Patan","Sabarkantha","Amreli","Rajkot","Junagadh","Bhavnagar","Jamnagar","Porbandar","Kachchh","Anand","Kheda","Surat","The Dangs","Tapi","Navsari","Vadodara","Bharuch","Narmada","Dahod","Panch Mahals","Valsad"}
                        },
new StateDistrict {
                        State = "DAMAN & DIU",
                        Districts = new List<string> {"Diu","Daman"}
                        },
new StateDistrict {
                        State = "DADRA & NAGAR HAVELI",
                        Districts = new List<string> {"Dadra & Nagar Haveli"}
                        },
new StateDistrict {
                        State = "HARYANA",
                        Districts = new List<string> {"Ambala","Yamuna Nagar","Panchkula","Bhiwani","Faridabad","Gurgaon","Rewari","Mahendragarh","Hisar","Sirsa","Fatehabad","Karnal","Panipat","Jind","Kaithal","Kurukshetra","Jhajjar","Rohtak","Sonipat"}
                        },
new StateDistrict {
                        State = "HIMACHAL PRADESH",
                        Districts = new List<string> {"Chamba","Kangra","Bilaspur (HP)","Hamirpur(HP)","Una","Mandi","Kullu","Lahul & Spiti","Kinnaur","Shimla","Sirmaur","Solan"}
                        },
new StateDistrict {
                        State = "JAMMU & KASHMIR",
                        Districts = new List<string> {"Bandipur","Baramulla","Kupwara","Jammu","Kathua","Udhampur","Poonch","Kargil","Leh","Rajauri","Reasi","Srinagar","Budgam","Ananthnag","Shopian","Pulwama","Kulgam","Doda"}
                        },
new StateDistrict {
                        State = "JHARKHAND",
                        Districts = new List<string> {"Dhanbad","Bokaro","Giridh","Hazaribag","Chatra","Ramgarh","Koderma","Latehar","Garhwa","Palamau","Ranchi","Gumla","Simdega","Lohardaga","West Singhbhum","Khunti","Deoghar","Godda","Jamtara","Sahibganj","Dumka","Pakur","Seraikela-kharsawan","East Singhbhum"}
                        },
new StateDistrict {
                        State = "KARNATAKA",
                        Districts = new List<string> {"Bangalore","Bangalore Rural","Ramanagar","Bagalkot","Bijapur(KAR)","Belgaum","Davangere","Bellary","Bidar","Dharwad","Gadag","Koppal","Yadgir","Gulbarga","Haveri","Uttara Kannada","Raichur","Chickmagalur","Chitradurga","Hassan","Kodagu","Chikkaballapur","Kolar","Mandya","Dakshina Kannada","Udupi","Mysore","Chamrajnagar","Shimoga","Tumkur"}
                        },
new StateDistrict {
                        State = "KERALA",
                        Districts = new List<string> {"Wayanad","Kozhikode","Malappuram","Kannur","Kasargod","Palakkad","Alappuzha","Ernakulam","Kottayam","Pathanamthitta","Idukki","Thrissur","Kollam","Thiruvananthapuram"}
                        },
new StateDistrict {
                        State = "LAKSHADWEEP",
                        Districts = new List<string> {"Lakshadweep"}
                        },
new StateDistrict {
                        State = "MADHYA PRADESH",
                        Districts = new List<string> {"Seoni","Balaghat","Mandla","Dindori","Bhopal","Raisen","Chhatarpur","Tikamgarh","Panna","Betul","Chhindwara","Hoshangabad","Narsinghpur","Harda","Satna","Rewa","Damoh","Sagar","Anuppur","Umaria","Shahdol","Sidhi","Singrauli","Vidisha","Ashok Nagar","Shivpuri","Guna","Gwalior","Datia","Bhind","Morena","Sheopur","Indore","Dewas","Dhar","Katni","Jabalpur","East Nimar","West Nimar","Barwani","Khandwa","Burhanpur","Khargone","Neemuch","Mandsaur","Jhabua","Ratlam","Alirajpur","Sehore","Rajgarh","Ujjain","Shajapur"}
                        },
new StateDistrict {
                        State = "MAHARASHTRA",
                        Districts = new List<string> {"Jalna","Aurangabad","Beed","Jalgaon","Dhule","Nandurbar","Nashik","Nanded","Latur","Osmanabad","Hingoli","Parbhani","Kolhapur","Ratnagiri","Sindhudurg","Satara","Sangli","Mumbai","Raigarh(MH)","Thane","Akola","Washim","Amravati","Buldhana","Gadchiroli","Chandrapur","Nagpur","Gondia","Bhandara","Wardha","Yavatmal","Ahmed Nagar","Solapur","Pune"}
                        },
new StateDistrict {
                        State = "GOA",
                        Districts = new List<string> {"South Goa","North Goa"}
                        },
new StateDistrict {
                        State = "MANIPUR",
                        Districts = new List<string> {"Imphal West","Churachandpur","Chandel","Thoubal","Tamenglong","Ukhrul","Imphal East","Bishnupur","Senapati"}
                        },
new StateDistrict {
                        State = "MIZORAM",
                        Districts = new List<string> {"Aizawl","Mammit","Lunglei","Kolasib","Lawngtlai","Champhai","Saiha","Serchhip"}
                        },
new StateDistrict {
                        State = "NAGALAND",
                        Districts = new List<string> {"Zunhebotto","Dimapur","Wokha","Phek","Mokokchung","Kiphire","Tuensang","Mon","Kohima","Peren","Longleng"}
                        },
new StateDistrict {
                        State = "TRIPURA",
                        Districts = new List<string> {"South Tripura","West Tripura","Dhalai","North Tripura"}
                        },
new StateDistrict {
                        State = "ARUNACHAL PRADESH",
                        Districts = new List<string> {"Lower Dibang Valley","East Siang","Dibang Valley","West Siang","Lohit","Papum Pare","Tawang","West Kameng","East Kameng","Lower Subansiri","Changlang","Tirap","Kurung Kumey","Upper Siang","Upper Subansiri"}
                        },
new StateDistrict {
                        State = "MEGHALAYA",
                        Districts = new List<string> {"West Garo Hills","East Garo Hills","Jaintia Hills","East Khasi Hills","South Garo Hills","Ri Bhoi","West Khasi Hills"}
                        },
new StateDistrict {
                        State = "ODISHA",
                        Districts = new List<string> {"Ganjam","Gajapati","Kalahandi","Nuapada","Koraput","Rayagada","Nabarangapur","Malkangiri","Kandhamal","Boudh","Baleswar","Bhadrak","Kendujhar","Khorda","Puri","Cuttack","Jajapur","Kendrapara","Jagatsinghapur","Mayurbhanj","Nayagarh","Balangir","Sonapur","Angul","Dhenkanal","Sambalpur","Bargarh","Jharsuguda","Debagarh","Sundergarh"}
                        },
new StateDistrict {
                        State = "CHANDIGARH",
                        Districts = new List<string> {"Chandigarh"}
                        },
new StateDistrict {
                        State = "PUNJAB",
                        Districts = new List<string> {"Ropar","Mohali","Rupnagar","Patiala","Ludhiana","Fatehgarh Sahib","Sangrur","Barnala","Amritsar","Tarn Taran","Bathinda","Mansa","Muktsar","Moga","Faridkot","Firozpur","Fazilka","Gurdaspur","Pathankot","Hoshiarpur","Nawanshahr","Jalandhar","Kapurthala"}
                        },
new StateDistrict {
                        State = "RAJASTHAN",
                        Districts = new List<string> {"Ajmer","Rajsamand","Bhilwara","Chittorgarh","Banswara","Dungarpur","Kota","Baran","Jhalawar","Bundi","Tonk","Udaipur","Alwar","Bharatpur","Dholpur","Jaipur","Dausa","Sawai Madhopur","Karauli","Barmer","Bikaner","Churu","Jhujhunu","Jodhpur","Jaisalmer","Nagaur","Pali","Sikar","Sirohi","Jalor","Ganganagar","Hanumangarh"}
                        },
new StateDistrict {
                        State = "TAMIL NADU",
                        Districts = new List<string> {"Chennai","Vellore","Tiruvannamalai","Kanchipuram","Tiruvallur","Villupuram","Cuddalore","Coimbatore","Dharmapuri","Salem","Erode","Karur","Namakkal","Krishnagiri","Nilgiris","Dindigul","Kanyakumari","Sivaganga","Ramanathapuram","Tuticorin","Tirunelveli","Madurai","Theni","Virudhunagar","Ariyalur","Tiruchirappalli","Pudukkottai","Tiruvarur","Thanjavur","Nagapattinam","Perambalur"}
                        },
new StateDistrict {
                        State = "UTTAR PRADESH",
                        Districts = new List<string> {"Agra","Aligarh","Hathras","Bulandshahr","Gautam Buddha Nagar","Etah","Firozabad","Etawah","Auraiya","Jhansi","Jalaun","Lalitpur","Mainpuri","Mathura","Azamgarh","Allahabad","Kaushambi","Ghazipur","Jaunpur","Sonbhadra","Mirzapur","Pratapgarh","Varanasi","Chandauli","Sant Ravidas Nagar","Pilibhit","Bareilly","Bijnor","Budaun","Hardoi","Kheri","Meerut","Bagpat","Moradabad","Jyotiba Phule Nagar","Rampur","Muzaffarnagar","Saharanpur","Shahjahanpur","Mau","Shrawasti","Bahraich","Ballia","Siddharthnagar","Sant Kabir Nagar","Basti","Deoria","Kushinagar","Gonda","Balrampur","Gorakhpur","Maharajganj","Banda","Chitrakoot","Mahoba","Hamirpur","Kannauj","Farrukhabad","Fatehpur","Kanpur Nagar","Unnao","Kanpur Dehat","Barabanki","Faizabad","Ambedkar Nagar","Ghaziabad","Lucknow","Raebareli","Sitapur","Sultanpur"}
                        },
new StateDistrict {
                        State = "UTTARAKHAND",
                        Districts = new List<string> {"Haridwar","Almora","Bageshwar","Chamoli","Rudraprayag","Dehradun","Udham Singh Nagar","Nainital","Champawat","Pauri Garhwal","Pithoragarh","Tehri Garhwal","Uttarkashi"}
                        },
new StateDistrict {
                        State = "WEST BENGAL",
                        Districts = new List<string> {"Kolkata","North 24 Parganas","South 24 Parganas","Birbhum","Murshidabad","Nadia","Cooch Behar","Jalpaiguri","Darjiling","Malda","South Dinajpur","North Dinajpur","Bardhaman","Bankura","West Midnapore","East Midnapore","Hooghly","Howrah","Medinipur","Puruliya"}
                        },
new StateDistrict {
                        State = "ANDAMAN & NICOBAR ISLANDS",
                        Districts = new List<string> {"South Andaman","North And Middle Andaman","Nicobar"}
                        },
new StateDistrict {
                        State = "SIKKIM",
                        Districts = new List<string> {"East Sikkim","West Sikkim","South Sikkim","North Sikkim"}
                        }

                };
            }
        }

        public  List<string> GetAllStates()
        {
            return AllStatesWithDistricts.Select(sd => sd.State.Replace(" ", "").ToLower()).ToList();
        }
        public List<string> GetDistrictsOfState(string state)
        {
            return AllStatesWithDistricts.Where(sd => sd.State.Replace(" ", "").ToLower() == state).SingleOrDefault().Districts.ConvertAll(x=>x.Replace(" ", "").ToLower());
        }
    }
}