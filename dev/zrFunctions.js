
/* ############################################################# 
Google API project: 
https://console.developers.google.com/iam-admin/settings/project?project=api-project-228132735377

Project name:	zrFunctions
Project ID:	api-project-228132735377
Project number:	228132735377

Taxes.gs
*/

function tester(){
  
  var ti = 83629.00;
  var td = [[40120.00, 0.2005, 
             43953.00, 0.2415, 
             70651.00, 0.3115, 
             80242.00, 0.3298, 
             83237.00, 0.3539, 
             87907.00, 0.3941, 
             136270.00, 0.4341, 
             514090.00, 0.4641, 
             100000000.00, 0.4953, 
             11138.00, 0.1500, 
             9670.00, 0.0500, 
             11138.00, 0.1500, 
             8211.00, 0.0500, 
             2426.00, 
             914.00, 
             6916.00, 0.1500, 
             4721.00, 0.0500, 
             2000.00, 0.1500, 
             1337.00, 0.0500
            ]];
  
  var ret = taxDue( ti, td, 2426.00, 914.00, 57, true, 1196.00 );
  
  var zoo = 10.00
}

function nonRefundableCredit(_taxdue, _credit){
  return (_taxdue >= _credit) ? _credit : _taxdue;
}

function taxDue( _taxableincome, _taxdata, _cpppaid, _eipaid, _age, _claimspousal, _spousetaxableincome ) {
  
  var taxdatarow = _taxdata[0];
  
  var taxrates = taxdatarow.slice(0,18);
  
  var taxdue = 0;
  var lowlimit = 0;
  for(var i=0; i<(taxrates.length)-1; i+=2){
    var highlimit = taxrates[i];
    var rate = taxrates[i+1];
    if(_taxableincome > highlimit){
      var taxable = highlimit - lowlimit;
      lowlimit = highlimit;
      var tax = taxable * rate;
      taxdue += tax;
    }
    else{
      var taxable = _taxableincome - lowlimit;
      var tax = taxable * rate;
      taxdue += tax;
      break;
    }
  }
  
  var fedpersamount = taxdatarow[18]; 
  var fedpersrate = taxdatarow[19]; 
  var fedperscred = nonRefundableCredit(taxdue, fedpersamount * fedpersrate);
  taxdue -= fedperscred
  
  var provpersamount = taxdatarow[20]; 
  var provpersrate = taxdatarow[21]; 
  var provperscred = nonRefundableCredit(taxdue, provpersamount * provpersrate);
  taxdue -= provperscred
  
  if(_claimspousal){
    var fedspousamount = taxdatarow[22]; 
    var fedspousrate = taxdatarow[23]; 
    var fedspouscred = nonRefundableCredit(taxdue, (fedspousamount-_spousetaxableincome) * fedspousrate);
    taxdue -= fedspouscred
  
    var provspousamount = taxdatarow[24]; 
    var provspousrate = taxdatarow[25]; 
    var provspouscred = nonRefundableCredit(taxdue, (provspousamount-_spousetaxableincome) * provspousrate);
    taxdue -= provspouscred
  }

  var cppmax = taxdatarow[26]; // Already subtracted from _taxableincome 
  if(_cpppaid > 0){
    _cpppaid = 0
  }
  if(-(_cpppaid) > cppmax){
    // Return CPP overpayment:
    var cppover = -(_cpppaid) - cppmax;
    taxdue -= cppover
  }

  var eimax = taxdatarow[27]; // Already subtracted from _taxableincome 
  if(_eipaid > 0){
    _eipaid = 0
  }
  if(-(_eipaid) > eimax){
    // Return EI overpayment:
    var eiover = -(_eipaid) - eimax;
    taxdue -= eiover
  }

  if (_age > 65){
    var fedageamount = taxdatarow[28];
    var fedagerate = taxdatarow[29];
    var fedagecred = nonRefundableCredit(taxdue, fedageamount * fedagerate);
    taxdue -= fedagecred
  
    var provageamount = taxdatarow[30];
    var provagerate = taxdatarow[31];
    var provagecred = nonRefundableCredit(taxdue, provageamount * provagerate);
    taxdue -= provagecred
  }
  
  //TODO: calculate application of these credits:
  var fedpincamount = taxdatarow[32]; // Off qualified pension income
  var fedpincrate = taxdatarow[33]; // Off qualified pension income
  
  //TODO: calculate application of these credits:
  var provpincamount = taxdatarow[34]; // Off qualified pension income
  var provpincrate = taxdatarow[35]; // Off qualified pension income
  
  //TODO: Pension splitting in retirement: http://www.cra-arc.gc.ca/pensionsplitting/
      
  return taxdue;
};
