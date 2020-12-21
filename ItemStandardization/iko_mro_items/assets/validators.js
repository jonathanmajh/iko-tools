const Database = require('./database')

class ManufacturerValidator {
    constructor() {
        this.CHAR_LIMIT = 30;
        this.db = new Database;
    }

    validateSingle(split_desc) {
        let manufacturer = false;
        // look from end of string since manufacturer name is mostly likely in last description
        for (let i=split_desc.length-1; i>=0; i--) {
            console.log('looking for manufacturer named: ' + split_desc[i]);
            manufacturer = this.db.isManufacturer(split_desc[i]);
            if(manufacturer) {
                console.log(manufacturer)
                return manufacturer
            }
        }
        return
    }
}

class Validate {
    validateSingle(raw_desc) {
        raw_desc = raw_desc.split(',');
        let split_desc = [];
        raw_desc.forEach(desc => {
            split_desc.push(desc.trim());
        });
        console.log(split_desc);
        let manuValid = new ManufacturerValidator();
        let manu = manuValid.validateSingle(split_desc);
        if (!manu) {
            console.log(`Warning: No Manufacturer Found for ${raw_desc}`);
        } else {
            split_desc = manu;
        }
        return(this.assembleDescription(split_desc));
    }

    async validateTriple(raw_desc) {
        // ['a', 'b', 'c']
        raw_desc = `${raw_desc[0]},${raw_desc[1]},${raw_desc[2]}`;
        return(this.validateSingle(raw_desc));
    }

    assembleDescription(split_desc) {
        let descriptions = ['', '', ''];
        for (let i = 0; i < split_desc.length - 1; i++) {
            split_desc[i] = `${split_desc[i]},`
        }
        split_desc.forEach((desc, j) => {
            if (j + 1 === split_desc.length) {
                descriptions[2] = `${descriptions[2]}${desc}`;
                return descriptions
            }
            descriptions.forEach((description, i) => {
                if (description.length + desc.length <= 30) {
                    descriptions[i] = `${description}${desc}`;
                }
            })
        })
    }
}

module.exports = Validate