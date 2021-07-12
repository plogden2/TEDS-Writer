import java.util.*;
import java.util.concurrent.TimeUnit;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.*;
import java.time.format.DateTimeFormatter;

import com.dalsemi.onewire.*;
import com.dalsemi.onewire.adapter.*;
import com.dalsemi.onewire.container.*;
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

                // if the device is an eprom
                if (owd.getDescription().indexOf("EPROM") != -1) {
                    eeprom_found = true;
                    System.out.println("============================");
                    System.out.println("Device Name: " + owd.getName());
                    System.out.println("Device Other Names: " + owd.getAlternateNames());
                    System.out.println("Device Description: " + owd.getDescription());
                    System.out.println("============================");

                    // set owd to max possible speed with available adapter, allow fall back
                    if (adapter.canOverdrive() && (owd.getMaxSpeed() == DSPortAdapter.SPEED_OVERDRIVE))
                        owd.setSpeed(owd.getMaxSpeed(), true);

                    if (args.length == 0 || (args[0].indexOf("r") == -1)) {

                        // print the TEDS data in the spreadsheet
                        printTEDSData();

                        // promt the user to verify the data
                        Scanner scanner = new Scanner(System.in); // Create a Scanner object
                        String ans = ""; // user input
                        System.out.println("============================");
                        System.out.println("Does this look correct (Y/N)?");
                        while (ans.length() == 0
                                || (ans.toLowerCase().charAt(0) != 'y' && ans.toLowerCase().charAt(0) != 'n'))
                            ans = scanner.nextLine();
                        scanner.close();

                        if (ans.toLowerCase().charAt(0) == 'y') {
                            clearEEPROM(owd);
                            byte[] buffer = writeTEDS(owd);
                            verifyTEDS(owd, buffer);
                        } else
                            System.out.println("Closing Application");
                    } else {
                        readEEPROM(owd);
                    }
                }
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
     * Erases all data from the EPROM
     *
     * @param device device to write to
     */
    public void clearEEPROM(OneWireContainer device) {
        System.out.println("\nErasing EEPROM");

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

                    System.out.println("Time to clear:\t" + Long.toString((end_time - start_time)) + "ms\n");

                } catch (Exception e) {
                    System.out.println("Exception in erasing: " + e + "  TRACE: ");
                    e.printStackTrace();
                }
            }
        }
        if (!found_bank)
            System.out.println("The device doesn't contain any memory banks");
    }

    /**
     * Writes the TEDS to the EPROM
     *
     * @param device device to write to
     * @return the byte array written
     */
    public byte[] writeTEDS(OneWireContainer device) {
        System.out.println("\nWriting data to EEPROM");

        byte[] buffer = {};// data array
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
                    formatData(buffer);

                    // start timer to time the dump of the bank contents
                    start_time = System.currentTimeMillis();

                    // write the buffer
                    bank.write(0, buffer, 0, bank.getSize());

                    end_time = System.currentTimeMillis();

                    System.out.println("Time to write:\t" + Long.toString((end_time - start_time)) + "ms\n");

                } catch (Exception e) {
                    System.out.println("Exception in writing: " + e + "  TRACE: ");
                    e.printStackTrace();
                }

            }
        }

        if (!found_bank)
            System.out.println("The device doesn't contain any memory banks");

        return buffer;
    }

    /**
     * Reads the EPROM and checks if the data matches the data sent
     *
     * @param device device to write to
     * @param buffer the byte array to compare to
     */
    public void verifyTEDS(OneWireContainer device, byte[] buffer) {
        System.out.println("\nVerifying data on EEPROM");

        byte[] read_buf = {};// data array
        boolean found_bank = false;
        int i, reps = 10;

        // loop through all of the memory banks on device
        // get the port names we can use and try to open, test and close each
        for (Enumeration bank_enum = device.getMemoryBanks(); bank_enum.hasMoreElements();) {

            // get the next memory bank
            MemoryBank bank = (MemoryBank) bank_enum.nextElement();
            if (bank.getBankDescription().toLowerCase().indexOf("main") != -1) {// if its the main memory

                // found a memory bank
                found_bank = true;

                try {
                    read_buf = new byte[bank.getSize()];

                    // get overdrive going so not a factor in time tests
                    bank.read(0, false, read_buf, 0, 1);

                    // dynamically change number of reps
                    reps = 1500 / read_buf.length;

                    if (device.getMaxSpeed() == DSPortAdapter.SPEED_OVERDRIVE)
                        reps *= 2;
                    // read the entire bank
                    for (i = 0; i < reps; i++)
                        bank.read(0, false, read_buf, 0, bank.getSize());

                    // compare to buffer
                    if (Arrays.equals(read_buf, buffer))
                        System.out.println("EEPROM Verification Successful");
                    else {
                        System.out.println("EEPROM Verification Failed\nData Mismatch\n");
                        System.out.println("Found Data:\n" + bytesToHex(read_buf));
                    }

                } catch (Exception e) {
                    System.out.println("Exception in verifying: " + e + "  TRACE: ");
                    e.printStackTrace();
                }

            }
        }

        if (!found_bank)
            System.out.println("The device doesn't contain any memory banks");
    }

    /**
     * Reads the EPROM and prints the contents
     *
     * @param device device to read
     */
    public void readEEPROM(OneWireContainer device) {
        System.out.println("\nReading data on EEPROM");

        byte[] read_buf = {};// data array
        boolean found_bank = false;
        int i, reps = 10;

        // loop through all of the memory banks on device
        // get the port names we can use and try to open, test and close each
        for (Enumeration bank_enum = device.getMemoryBanks(); bank_enum.hasMoreElements();) {

            // get the next memory bank
            MemoryBank bank = (MemoryBank) bank_enum.nextElement();
            if (bank.getBankDescription().toLowerCase().indexOf("main") != -1) {// if its the main memory

                // found a memory bank
                found_bank = true;

                try {
                    read_buf = new byte[bank.getSize()];

                    // get overdrive going so not a factor in time tests
                    bank.read(0, false, read_buf, 0, 1);

                    // dynamically change number of reps
                    reps = 1500 / read_buf.length;

                    if (device.getMaxSpeed() == DSPortAdapter.SPEED_OVERDRIVE)
                        reps *= 2;
                    // read the entire bank
                    for (i = 0; i < reps; i++)
                        bank.read(0, false, read_buf, 0, bank.getSize());

                    System.out.println("Contents:");
                    System.out.println(bytesToHex(read_buf));

                } catch (Exception e) {
                    System.out.println("Exception in reading: " + e + "  TRACE: ");
                    e.printStackTrace();
                }

            }
        }

        if (!found_bank)
            System.out.println("The device doesn't contain any memory banks");
    }

    /**
     * formats the buffer to fill usable memory
     *
     * @param buffer byte array to add data to
     */
    public void formatBuffer(byte[] buffer) throws Exception {
        Arrays.fill(buffer, (byte) 0);
        // fill usable portion of memory bank
        for (int i = 0; i < 8 * 32; i++) {
            buffer[i] = (byte) ((i % 32 == 0) ? 31 : 255);
        }
    }

    /**
     * formats the data into a byte array to be written
     *
     * @param buffer byte array to add data to
     */
    public void formatData(byte[] buffer) throws Exception {
        String str_entry, type;
        char char_entry;
        int int_entry = 0;
        float float_entry = 0;
        int length = 0;
        int byte_length = 0;
        byte[] byte_arr = { (byte) 0, (byte) 0, (byte) 0, (byte) 0 };

        // for each value in data map
        for (String key : map_keys) {

            // get data
            type = data_map.get(key)[3].replaceAll("\\s", "");
            length = (int) Double.parseDouble(data_map.get(key)[1]);
            if (type.equals("UNINT"))
                int_entry = (int) Double.parseDouble(data_map.get(key)[0]);
            else if (type.equals("Chr5")) {
                str_entry = data_map.get(key)[0];
                int_entry = 0;
                for (int i = 0; i < str_entry.length(); i++) {
                    char_entry = str_entry.charAt(i);
                    int_entry += (int) ((((byte) char_entry << 3) >> 3) - 64) << 5 * i;
                }
            } else if (type.equals("DATE")) {
                LocalDate date_1 = LocalDate.parse("1998-01-01");
                Instant instant_1 = date_1.atStartOfDay(ZoneId.systemDefault()).toInstant();
                LocalDate date_2 = LocalDate.parse(data_map.get(key)[0]);
                Instant instant_2 = date_2.atStartOfDay(ZoneId.systemDefault()).toInstant();

                int_entry = (int) Duration.between(instant_1, instant_2).toDays();

            } else if (type.equals("ConRelRes")) {
                int_entry = conRelResToInt(data_map.get(key)[0], data_map.get(key)[2]);
            } else if (type.equals("ConRes")) {
                int_entry = conResToInt(data_map.get(key)[0], data_map.get(key)[2]);
            } else if (type.equals("Single")) {
                float_entry = (float) Double.parseDouble(data_map.get(key)[0]);
                int_entry = Float.floatToIntBits(float_entry);
            } else
                throw new Exception("Unrecognized data type " + data_map.get(key)[3]);

            System.out.println(key + ": Pos: " + (sub_index == 0 ? index * 8 : ((index - 1) * 8 + sub_index)));// print
                                                                                                               // bitpos
                                                                                                               // for
                                                                                                               // each
                                                                                                               // element

            // put data into byte array
            byte_arr[0] = (byte) (int_entry >>> 24);
            byte_arr[1] = (byte) (int_entry >>> 16);
            byte_arr[2] = (byte) (int_entry >>> 8);
            byte_arr[3] = (byte) ((int_entry < 0) ? int_entry + 256 : int_entry);

            byte_length = length;
            if (byte_length > 0)
                addToBuffer(buffer, byte_arr[3], byte_length);
            byte_length = length - 8;
            if (byte_length > 0)
                addToBuffer(buffer, byte_arr[2], byte_length);
            byte_length = length - 16;
            if (byte_length > 0)
                addToBuffer(buffer, byte_arr[1], byte_length);
            byte_length = length - 24;
            if (byte_length > 0)
                addToBuffer(buffer, byte_arr[0], byte_length);

            int_entry = 0;
        }
        sub_index = 0;
        addToBuffer(buffer, (byte) 255, 8);

        calculateChecksum(buffer);

        System.out.println(bytesToHex(buffer));
    }

    /**
     * Add at most a byte to the buffer filling in empty space.
     * 
     * @param buffer byte array to add to
     * @param data   byte to add to the array
     * @param length length of the data
     */
    public void addToBuffer(byte[] buffer, byte data, int length) {
        if (index % 32 == 0)// skip checksum bytes
            index++;
        if (length > 8)
            length = 8;

        if (sub_index == 0) {
            buffer[index] = (byte) (data << 8 - (length));// shift the data all the way to the left in an empty byte
            sub_index += length;
            if (sub_index >= 8) {
                sub_index = 0;
                index++;// move to the next byte
                if (index % 32 == 0)// skip checksum bytes
                    index++;
            }
        } else {
            int temp = 8 - sub_index;
            int temp_byte = buffer[index];
            if (temp_byte < 0)
                temp_byte += 256;
            buffer[index] = (byte) ((byte) (temp_byte >> temp) + (byte) (data << sub_index));// move existing buffer
                                                                                             // data
                                                                                             // right, shift new data
                                                                                             // left,
                                                                                             // and add new data
            sub_index += length;
            if (sub_index >= 8) {
                sub_index = 0;
                index++;// move to the next byte
                if (index % 32 == 0)// skip checksum bytes
                    index++;
                data = (byte) (data >> temp);// remove the bits that were added to the
                buffer[index] = (byte) (data << 8 - (length - temp));// shift the leftover data all the way to the left
                                                                     // in an
                sub_index += length - temp;
            } else {
                buffer[index] = (byte) (buffer[index] << (8 - sub_index)); // shift data left in not full
            }
        }
    }

    /**
     * Calculates and inserts the checksum at the beginning of the byte array.
     * 
     * @param buffer byte array
     */
    public void calculateChecksum(byte[] buffer) {
        for (int block = 0; block <= (int) index / 32; block++) {
            System.out.println(block);
            int sum = 0;
            for (int i = 1; i < 32; ++i) {
                sum += buffer[block * 32 + i];
            }
            // modulus 256 sum
            sum %= 256;
            // twos complement
            byte twos_complement = (byte) (~(sum) + 1);
            buffer[block * 32] = twos_complement;
        }
    }

    /**
     * Gets the TEDS Data and then prints the TEDS data to the console
     */
    public void printTEDSData() throws Exception {
        getTEDSData();
        System.out.println("Data to be Written:\n");
        for (String key : map_keys) {
            String value = data_map.get(key)[0].toString();
            System.out.println(String.format("%30s", key) + ":\t" + String.format("%-20s", value));
        }
    }

    /**
     * Reads the TEDS data from the xlsx file
     */
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
     * Converts a byte array to a string of hexadecimal values.
     * 
     * @param bytes the byte array being converted
     * @return a string of hexidecimal values
     */
    public static String bytesToHex(byte[] bytes) {
        final byte[] HEX_ARRAY = "0123456789ABCDEF".getBytes(StandardCharsets.US_ASCII);
        byte[] hexChars = new byte[bytes.length * 2];
        for (int j = 0; j < bytes.length; j++) {
            int v = bytes[j] & 0xFF;
            hexChars[j * 2] = HEX_ARRAY[v >>> 4];
            hexChars[j * 2 + 1] = HEX_ARRAY[v & 0x0F];
        }
        return new String(hexChars, StandardCharsets.UTF_8);
    }

    /**
     * Converts conRes data to an int.
     * 
     * @param entry the vaulue being evaluated
     * @param range the string representation of the specified range
     */
    public int conResToInt(String entry, String range) {
        int value = 0;
        double double_entry = Double.valueOf(entry);
        double min = Double.valueOf(range.substring(0, range.indexOf("to")));
        double step = Double.valueOf(range.substring(range.indexOf("step") + 5));

        value = (int) Math.round((double_entry - min) / step);
        System.out.println(value);

        return value;
    }

    /**
     * Converts conRelRes data to an int.
     * 
     * @param entry the vaulue being evaluated
     * @param range the string representation of the specified range
     */
    public int conRelResToInt(String entry, String range) {
        int value = 0;
        double double_entry = Double.valueOf(entry);
        double min = Double.valueOf(range.substring(0, range.indexOf("to")));
        double resolution = Double.valueOf(range.substring(range.indexOf("±") + 1, range.indexOf("%"))) / 100;

        value = (int) Math.round(((1 / (Math.log10(1 + resolution * 2))) * Math.log10(double_entry / min)));

        System.out.println(value);
        return value;
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
