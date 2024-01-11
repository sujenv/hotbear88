from commonmd import SubWindowBase

# For First Db
from customer import CustomerDialog
from customercar import CustomerCarDialog
from customer_regno import CustomerRegNoDialog
from employee import EmployeeDialog
from employeecar import EmployeeCarDialog
from employee_regno import EmployeeRegNoDialog
from employee_fone import EmployeePhoneNoDialog
from employee_address import EmployeeAddressDialog
from customer_fone import CustomerPhoneNoDialog
from customer_address import CustomerAddressDialog
from customer_bkacc import CustomerBkAccDialog

from delivery import DeliveryDialog
from salesitem import SalesItemDialog
from salesprice import SalesPriceDialog
from product import ApProductDialog
from cost import CostDialog
from advancepay import AdvancePayDialog
from bankacc_product_info import BankAccProductInfoDialog
from bankacc_product_gen import BankAccProductGenDialog
from bankacc_product_chk import BankAccProductChkDialog
from gpai_payment_info import GPAIInfoDialog
from gpai_payment_gen import GPAIGenDialog
from gpai_payment_view import GPAIViewDialog
from payment_product import PaymentProductDialog
from oil_payment_info import OilUsageInfoDialog
from oil_payment_gen import OilUsageGenDialog
from oil_payment_view import OilUsageViewDialog
from oil_paymentemp_view import OilUsageEmpViewDialog
from support_apt_info import SupportAptInfoDialog
from support_apt_gen import SupportAptGenDialog
from support_apt_view import SupportAptEmpViewDialog


# For Second Db
from consumableproduct import ConsumableProductDialog
from consumableinprice import ConsumableInPriceDialog
from consumableoutprice import ConsumableOutPriceDialog
from conversion import ConversionDialog
from consumablereceipt import ConsumableReceiptDialog
from consumablesales import ConsumableSalesDialog
from inventory import CurrentInventoryDialog
from consumableinclosing import ConsumableInClosingDialog
from consumableoutclosing import ConsumableOutClosingDialog

# For Third Db
from calmaster import CalMasterDialog
from calwkhrinhot import CalWorkingHrInhOtDialog
from calwkhrouthot import CalWorkingHrOuthOtDialog
from calpension import CalPensionDialog
from calemployeeinfo import CalEmployeeInfoDialog
from calc_absenteeism import CalAbsenteeismDialog
from calc_inh_salary import CalSalaryInhDialog
from calc_outh_salary import CalSalaryOuthDialog

# For Fourth Db
from aptmaster import AptMasterDialog
from aptcontact import AptContactDialog
from aptaddress import AptAddressDialog
from aptcontract import AptContractDialog
from aptcontractpic import AptContractPicDialog
from aptreport import AptReportDialog
from reinfo import RecyclingDialog
from aptclothbtb import AptClothBtbDialog


# Define the user_credentials dictionary as a class attribute
userlist = {"이종욱": "6074", "이지아": "5521", "주혜진": "5506"}