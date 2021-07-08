import java.util.*;

import javax.lang.model.util.ElementScanner14;

import java.io.*;
import java.nio.charset.StandardCharsets;

import com.dalsemi.onewire.*;
import com.dalsemi.onewire.adapter.*;
import com.dalsemi.onewire.container.*;
import com.dalsemi.onewire.utils.*;
import org.apache.poi.ss.usermodel.*;

/**
 * TEDS Writer Tool
 *
 * @version 1.0, 7 July 2021
 * @author Patrick Ogden
 */
public class TEDS_Writer {
    private HashMap<String, String[]> data_map = new HashMap<>();
    private ArrayList<String> map_keys = new ArrayList<>();
    private int index = 0;// current byte of buffer
    private int sub_index = 0;// number of bits filled in byte

    public TEDS_Writer(String[] args) throws Exception {
        try {
            boolean eeprom_found = false;
            // get the default adapter
            DSPortAdapter adapter = OneWireAccessProvider.getDefaultAdapter();

            // check if adapter can program EPROMS
            // if(adapter.canProgram())
            // System.out.println("Adapter Connected Successfully.");
            // else{
            // System.out.println("The connected adapter doesn't have EPROM programming
            // capabilities.");
            // throw new Exception("Adapter doesn't have EPROM programming capabilities");
            // }

            // get exclusive use of adapter
            adapter.beginExclusive(true);

            // clear any previous search restrictions
            adapter.setSearchAllDevices();
            adapter.targetAllFamilies();
            adapter.setSpeed(adapter.SPEED_REGULAR);

            // enumerate through all the iButtons found
            for (Enumeration owd_enum = adapter.getAllDeviceContainers(); owd_enum.hasMoreElements();) {

                // get the next owd
                OneWireContainer owd = (OneWireContainer) owd_enum.nextElement();

                printTEDSData();// remove after testing
                writeTEDS(owd);// remove after testing

                // if the device is an eprom
                // if (owd.getDescription().indexOf("EPROM") != -1) {
                // eeprom_found = true;
                // System.out.println("Device Name: " + owd.getName());
                // System.out.println("Device Other Names: " + owd.getAlternateNames());
                // System.out.println("Device Description: " + owd.getDescription());

                // // set owd to max possible speed with available adapter, allow fall back
                // if (adapter.canOverdrive() && (owd.getMaxSpeed() ==
                // DSPortAdapter.SPEED_OVERDRIVE))
                // owd.setSpeed(owd.getMaxSpeed(), true);

                // // print the TEDS data in the spreadsheet
                // printTEDSData();

                // // promt the user to verify the data
                // Scanner scanner = new Scanner(System.in); // Create a Scanner object
                // String ans = ""; // user input
                // System.out.println("Does this look correct (Y/N)?");
                // while (ans.length() == 0
                // || (ans.toLowerCase().charAt(0) != 'y' && ans.toLowerCase().charAt(0) !=
                // 'n'))
                // ans = scanner.nextLine();
                // scanner.close();

                // if (ans.toLowerCase().charAt(0) == 'y') {
                // clearEEPROM(owd);
                // writeTEDS(owd);
                // } else
                // System.out.println("Closing Application");
                // }
            }

            // end exclusive use of adapter
            adapter.endExclusive();

            if (!eeprom_found)
                System.out.println("No EEPROM Found");

            // free the port used by the adapter
            System.out.println("Releasing adapter port");
            adapter.freePort();
        } catch (Exception e) {
            System.out.println("Exception: " + e);
            e.printStackTrace();
        }

        System.exit(0);
    }

    /**
     * Writes the TEDS to the EPROM
     *
     * @param device device to write to
     */
    public void writeTEDS(OneWireContainer device) {
        System.out.println("Writing data to EEPROM");

        byte[] buffer;// data array
        long start_time, end_time;
        boolean found_bank = false;

        // loop through all of the memory banks on device
        // get the port names we can use and try to open, test and close each
        for (Enumeration bank_enum = device.getMemoryBanks(); bank_enum.hasMoreElements();) {

            // get the next memory bank
            MemoryBank bank = (MemoryBank) bank_enum.nextElement();
            if (bank.getBankDescription().toLowerCase().indexOf("main") != -1) {// if its the main memory

                // found a memory bank
                found_bank = true;

                try {
                    buffer = new byte[bank.getSize()];
                    formatBuffer(buffer);

                    // start timer to time the dump of the bank contents
                    start_time = System.currentTimeMillis();

                    // write the buffer
                    bank.write(0, buffer, 0, bank.getSize());

                    end_time = System.currentTimeMillis();

                    System.out.println("Time to write:\t" + Long.toString((end_time - start_time)) + "ms");

                } catch (Exception e) {
                    System.out.println("Exception in writing: " + e + "  TRACE: ");
                    e.printStackTrace();
                }

            }
        }

        if (!found_bank)
            System.out.println("The device doesn't contain any memory banks");
    }

    /**
     * Erases all data from the EPROM
     *
     * @param device device to write to
     */
    public void clearEEPROM(OneWireContainer device) {
        System.out.println("Erasing EEPROM");

        byte[] buffer;// data array
        long start_time, end_time;
        boolean found_bank = false;

        // loop through all of the memory banks on device
        // get the port names we can use and try to open, test and close each
        for (Enumeration bank_enum = device.getMemoryBanks(); bank_enum.hasMoreElements();) {

            // get the next memory bank
            MemoryBank bank = (MemoryBank) bank_enum.nextElement();

            // found a memory bank
            found_bank = true;

            try {
                buffer = new byte[bank.getSize()];
                Arrays.fill(buffer, (byte) 0);

                // start timer to time the dump of the bank contents
                start_time = System.currentTimeMillis();

                // write the buffer
                bank.write(0, buffer, 0, bank.getSize());

                end_time = System.currentTimeMillis();

                System.out.println("Time to clear:\t" + Long.toString((end_time - start_time)) + "ms");

            } catch (Exception e) {
                System.out.println("Exception in erasing: " + e + "  TRACE: ");
                e.printStackTrace();
            }
        }

        if (!found_bank)
            System.out.println("The device doesn't contain any memory banks");
    }

    /**
     * formats the data into a byte array to be written
     *
     * @param buffer byte array to add data to
     */
    public void formatBuffer(byte[] buffer) throws Exception {
        String str_entry, type;
        char char_entry;
        int int_entry = 0;
        int length = 0;
        int byte_length = 0;
        byte[] byte_arr = { (byte) 0, (byte) 0, (byte) 0, (byte) 0 };

        Arrays.fill(buffer, (byte) 0);

        // for each value in data map
        for (String key : map_keys) {
            // get data
            type = data_map.get(key)[3].replaceAll("\\s", "");
            length = (int) Double.parseDouble(data_map.get(key)[1]);
            if (type.equals("UNINT") || type.equals("ConRes") || type.equals("ConRelRes"))
                int_entry = (int) Double.parseDouble(data_map.get(key)[0]);
            else if (type.equals("Chr5")) {
                str_entry = data_map.get(key)[0];
                int_entry = 0;
                for (int i = 0; i < str_entry.length(); i++) {
                    char_entry = str_entry.charAt(i);
                    int_entry += (int) ((((byte) char_entry << 3) >> 3) - 64) << 5 * i;
                }

            } else
                throw new Exception("Unrecognized data type " + data_map.get(key)[3]);

            // put data into byte array
            byte_arr[0] = (byte) (int_entry >>> 24);
            byte_arr[1] = (byte) (int_entry >>> 16);
            byte_arr[2] = (byte) (int_entry >>> 8);
            byte_arr[3] = (byte) ((int_entry < 0) ? int_entry + 256 : int_entry);

            byte_length = length;
            if (!(byte_length <= 0))
                addToBuffer(buffer, byte_arr[3], byte_length);
            byte_length = length - 8;
            if (!(byte_length <= 0))
                addToBuffer(buffer, byte_arr[2], byte_length);
            byte_length = length - 16;
            if (!(byte_length <= 0))
                addToBuffer(buffer, byte_arr[1], byte_length);
            byte_length = length - 24;
            if (!(byte_length <= 0))
                addToBuffer(buffer, byte_arr[0], byte_length);

        }

        calculateChecksum(buffer);

        System.out.println(bytesToHex(buffer));
    }

    // add at most a byte to the buffer filling in empty space
    public void addToBuffer(byte[] buffer, byte data, int length) {
        if (index == 0)// skip checksum byte
            index++;
        if (length > 8)
            length = 8;

        if (sub_index == 0) {
            buffer[index] = (byte) (data << 8-(length));// shift the data all the way to the left in an empty byte
            sub_index += length;
            if (sub_index >= 8) {
                sub_index = 0;
                index++;// move to the next byte
            }
        } else {
            int temp = 8 - sub_index;
            int temp_byte = buffer[index];
            if(temp_byte<0)temp_byte+=256;
            buffer[index] = (byte) ((byte)(temp_byte>>temp) + (byte)(data << sub_index));// move existing buffer data
                                                                                          // right, shift new data left,
                                                                                          // and add new data
            
            if (sub_index >= 8) {
                sub_index = 0;
                index++;// move to the next byte
                data = (byte) (data >> temp);// remove the bits that were added to the
                buffer[index] = (byte) (data << 8-(length-temp));// shift the leftover data all the way to the left in an
                sub_index += length - temp;
            } else {
                buffer[index] = (byte) (buffer[index] << (8 - sub_index)); // shift data left in not full
            }
        }
    }

    public void calculateChecksum(byte[] buffer) {
        int sum = 0;
        for (int i = 0; i < (buffer.length - 1); ++i) {
            sum += buffer[i];
        }
        // modulo 256 sum
        sum %= 256;

        // twos complement
        byte twos_complement = (byte) (~(sum) + 1);

        buffer[0] = twos_complement;
    }

    private static final byte[] HEX_ARRAY = "0123456789ABCDEF".getBytes(StandardCharsets.US_ASCII);

    public static String bytesToHex(byte[] bytes) {
        byte[] hexChars = new byte[bytes.length * 2];
        for (int j = 0; j < bytes.length; j++) {
            int v = bytes[j] & 0xFF;
            hexChars[j * 2] = HEX_ARRAY[v >>> 4];
            hexChars[j * 2 + 1] = HEX_ARRAY[v & 0x0F];
        }
        return new String(hexChars, StandardCharsets.UTF_8);
    }

    // Prints the TEDS data to the console
    public void printTEDSData() throws Exception {
        getTEDSData();
        System.out.println("Data to be Written:");
        for (String key : map_keys) {
            String value = data_map.get(key)[0].toString();
            System.out.println(key + ":\t" + value);
        }
    }

    // Reads the TEDS data from the xlsx file
    public void getTEDSData() throws Exception {
        if (map_keys.size() == 0) {// only get data once
            Workbook workbook = WorkbookFactory.create(new FileInputStream("TEDS_Data.xlsx")); // open XLSX workbook
            Sheet sheet = workbook.getSheetAt(0); // set to first sheet
            Iterator<Row> row_iterator = sheet.iterator(); // create a row iterator
            Row row = row_iterator.next();
            Cell cell;
            String field, length, range, type, entry;
            String[] arr;
            while (row_iterator.hasNext()) {
                row = row_iterator.next(); // For each row, iterate through each columns
                if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.STRING) {
                    // add field and entry to hash map
                    cell = row.getCell(0);
                    field = cell.getStringCellValue();
                    cell = row.getCell(1);
                    length = (cell.getCellType() == CellType.STRING) ? cell.getStringCellValue()
                            : String.valueOf(cell.getNumericCellValue());
                    cell = row.getCell(2);
                    range = (cell.getCellType() == CellType.STRING) ? cell.getStringCellValue()
                            : String.valueOf(cell.getNumericCellValue());
                    cell = row.getCell(3);
                    type = (cell.getCellType() == CellType.STRING) ? cell.getStringCellValue()
                            : String.valueOf(cell.getNumericCellValue());
                    cell = row.getCell(4);
                    entry = (cell.getCellType() == CellType.STRING) ? cell.getStringCellValue()
                            : String.valueOf(cell.getNumericCellValue());
                    if (withinRange(entry, range)) {
                        arr = new String[] { entry, length, range, type };
                        data_map.put(field, arr);
                        map_keys.add(field);
                    } else
                        throw new Exception(entry + " is outside of the range for " + field);
                }
            }
            workbook.close();
        }
    }

    /**
     * Checks if the entry is within the given range and returns the result.
     * 
     * @param entry the vaulue being evaluated
     * @param range the string representation of the specified range
     */
    public boolean withinRange(String entry, String range) {
        // TODO check if entry is within range
        return true;
    }
}
