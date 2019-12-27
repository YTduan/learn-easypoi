package com.example.easypoi.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;

import java.io.Serializable;
import java.util.Date;

/**
 * @author duanyuantong
 * @version Id: People, v 0.1 2019-11-20 17:27 duanyuantong Exp $
 */
public class People implements Serializable {

    /**姓名*/
    @Excel(name = "姓名")
    private String  name;

    /**性别*/
    @Excel(name = "性别")
    private Integer sex;

    /** 年龄*/
    @Excel(name = "年龄")
    private Integer age;

    /**体重*/
    @Excel(name = "体重")
    private double  weight;

    /** 身高*/
    @Excel(name = "身高")
    private double  height;

    @Excel(name="照片")
    private byte[] pic;

    /** 生日*/
    @Excel(name = "生日",databaseFormat = "yyyyMMddHHmmss", format = "yyyy-MM-dd")
    private Date    birthday;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getSex() {
        return sex;
    }

    public void setSex(Integer sex) {
        this.sex = sex;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public double getWeight() {
        return weight;
    }

    public void setWeight(double weight) {
        this.weight = weight;
    }

    public double getHeight() {
        return height;
    }

    public void setHeight(double height) {
        this.height = height;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public byte[] getPic() {
        return pic;
    }

    public void setPic(byte[] pic) {
        this.pic = pic;
    }

    public People() {
    }

    public People(String name, Integer sex, Integer age, double weight, double height, Date birthday) {
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.weight = weight;
        this.height = height;
        this.birthday = birthday;
    }

    public People(String name) {
        this.name = name;
    }
}
